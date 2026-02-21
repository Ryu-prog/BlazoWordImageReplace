using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using SixLabors.ImageSharp; // 完全マネージド(C#)
using SixLabors.ImageSharp.PixelFormats;

namespace WordImageReplace
{
    /// <summary>
    /// Open XML を使用して、Word ドキュメント テンプレート内のヘッダー イメージを置換または追加する機能を提供します。
    /// </summary>
    /// <remarks>この静的クラスを使用すると、呼び出し元に許可を与えることで Word 文書のヘッダー画像を変更できます。</remarks>
    public static class OpenXmlWordHeaderReplacer
    {
        const long EMU_PER_PIXEL = 9525;

        public static byte[] ReplaceHeaderImages(
            byte[] templateBytes,
            byte[] frontImageBytes,
            byte[] backImageBytes,
            bool changeFront,
            bool changeBack)
        {
            // 1. テンプレートをメモリストリームに読み込む
            using var memStream = new MemoryStream();
            memStream.Write(templateBytes, 0, templateBytes.Length);

            using (var doc = WordprocessingDocument.Open(memStream, true))
            {
                var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException();
                var sectionProps = mainPart.Document.Descendants<SectionProperties>().ToList();

                foreach (var sec in sectionProps)
                {
                    foreach (var headerRef in sec.Elements<HeaderReference>())
                    {
                        var type = headerRef.Type?.Value ?? HeaderFooterValues.Default;
                         var headerPart = (HeaderPart)mainPart.GetPartById(headerRef.Id.Value);

                if (type == HeaderFooterValues.Default && changeFront)
                            ReplaceOrAddImageInHeader(headerPart, frontImageBytes);
                        else if (type == HeaderFooterValues.Even && changeBack)
                            ReplaceOrAddImageInHeader(headerPart, backImageBytes);
                    }
                }
            }
            return memStream.ToArray(); // 処理後のファイルをバイト配列で返す
        }

        /// <summary>
        /// 指定されたヘッダー部分の既存の画像を新しい画像に置き換えます。画像が存在しない場合は新しい画像を追加します。
        /// </summary>
        /// <param name="headerPart"></param>
        /// <param name="imageBytes"></param>
        private static void ReplaceOrAddImageInHeader(HeaderPart headerPart, byte[] imageBytes)
        {
            if (imageBytes == null || imageBytes.Length == 0) return;

            // 1. ImageSharpでサイズ取得
            long cx = 0, cy = 0;
            using (var image = Image.Load<Rgba32>(imageBytes))
            {
                cx = (long)(image.Width * 9525);
                cy = (long)(image.Height * 9525);
            }

            var existingImagePart = headerPart.ImageParts.FirstOrDefault();

            if (existingImagePart != null)
            {
                // --- 修正ポイント A: 既存画像がある場合 ---
                // 画像バイナリを上書きするだけでOKです。
                // Word側ですでに画像が配置されているため、XML構造（Drawing要素）の追加は不要です。
                using (var partStream = existingImagePart.GetStream(FileMode.Create, FileAccess.Write))
                {
                    partStream.Write(imageBytes, 0, imageBytes.Length);
                }

                // XML側でサイズ(Extent)が固定されている場合があるため、
                // 厳密には既存の Drawing 要素を探して cx/cy を書き換えるのがベストですが、
                // まずはこの「Appendしない」修正でファイル破損は回避できます。
                return;
            }
            else
            {
                // --- 修正ポイント B: 新規追加の場合 ---
                // 画像が全くない場合のみ、新規パート作成と XML(Drawing) の追加を行います。
                var imagePart = headerPart.AddImagePart(ImagePartType.Png);
                using (var imgStream = new MemoryStream(imageBytes))
                {
                    imagePart.FeedData(imgStream);
                }

                string rId = headerPart.GetIdOfPart(imagePart);
                var element = CreateImageDrawing(rId, cx, cy);

                if (headerPart.Header == null)
                {
                    headerPart.Header = new Header();
                }

                // 新規の時だけ Append する
                var paragraph = new Paragraph(new Run(element));
                headerPart.Header.Append(paragraph);
                headerPart.Header.Save();
            }
        }

        private static Drawing CreateImageDrawing(string relationshipId, long cx, long cy)
        {
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = cx, Cy = cy },
                        new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value)BitConverter.ToUInt32(Guid.NewGuid().ToByteArray(), 0),
                            Name = "Picture " + relationshipId
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)BitConverter.ToUInt32(Guid.NewGuid().ToByteArray(), 0),
                                            Name = "Image " + relationshipId
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip() { Embed = relationshipId },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = cx, Cy = cy }),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                                )
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U,
                        // behind の配置（必要ならさらに wrap を調整）
                    });

            return element;
        }
    }
}