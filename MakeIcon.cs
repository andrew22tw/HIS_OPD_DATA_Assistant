using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;

class MakeIcon {
    static void Main() {
        // Generate multiple sizes for .ico
        var sizes = new[]{16,32,48,256};
        var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        // ICO header
        bw.Write((short)0); bw.Write((short)1); bw.Write((short)sizes.Length);
        // Placeholder directory entries
        long dirStart = ms.Position;
        for(int i=0;i<sizes.Length;i++) {
            bw.Write(new byte[16]);
        }
        var offsets = new long[sizes.Length];
        var imgData = new byte[sizes.Length][];

        for(int si=0;si<sizes.Length;si++) {
            int sz = sizes[si];
            var bmp = DrawIcon(sz);
            var pngMs = new MemoryStream();
            bmp.Save(pngMs, System.Drawing.Imaging.ImageFormat.Png);
            imgData[si] = pngMs.ToArray();
            offsets[si] = ms.Position;
            bw.Write(imgData[si]);
            bmp.Dispose();
        }

        // Go back and fill directory entries
        ms.Position = dirStart;
        for(int i=0;i<sizes.Length;i++) {
            bw.Write((byte)(sizes[i]<256?sizes[i]:0)); // width
            bw.Write((byte)(sizes[i]<256?sizes[i]:0)); // height
            bw.Write((byte)0); // color palette
            bw.Write((byte)0); // reserved
            bw.Write((short)1); // color planes
            bw.Write((short)32); // bits per pixel
            bw.Write((int)imgData[i].Length); // size
            bw.Write((int)offsets[i]); // offset
        }

        File.WriteAllBytes("app.ico", ms.ToArray());
        Console.WriteLine("OK: app.ico created");
    }

    static Bitmap DrawIcon(int sz) {
        var bmp = new Bitmap(sz, sz);
        using(var g = Graphics.FromImage(bmp)) {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.Clear(Color.Transparent);

            float s = sz / 256f; // scale factor

            // === Background: rounded rectangle with gradient ===
            var bgRect = new RectangleF(sz*0.02f, sz*0.02f, sz*0.96f, sz*0.96f);
            var bgPath = RoundedRect(bgRect, sz*0.15f);
            using(var bgBrush = new LinearGradientBrush(bgRect,
                Color.FromArgb(34,139,230), Color.FromArgb(56,178,172), 45f))
                g.FillPath(bgBrush, bgPath);

            // === Document (back, slightly offset) ===
            float docL = sz*0.12f, docT = sz*0.18f, docW = sz*0.52f, docH = sz*0.68f;
            var docPath = RoundedRect(new RectangleF(docL, docT, docW, docH), sz*0.04f);
            using(var b = new SolidBrush(Color.FromArgb(180,255,255,255)))
                g.FillPath(b, docPath);

            // Document lines
            float lineX = docL + sz*0.06f;
            float lineW = docW - sz*0.12f;
            using(var p = new Pen(Color.FromArgb(120,34,139,230), Math.Max(sz*0.02f,1f))) {
                for(int i=0;i<4;i++) {
                    float ly = docT + sz*0.14f + i*sz*0.11f;
                    float lw = (i==3) ? lineW*0.6f : lineW;
                    g.DrawLine(p, lineX, ly, lineX+lw, ly);
                }
            }

            // === Stethoscope (front, overlapping document) ===
            // Earpieces (top Y shape)
            float sthCx = sz*0.62f, sthTop = sz*0.10f;
            using(var p = new Pen(Color.White, Math.Max(sz*0.035f, 1.5f))) {
                p.StartCap = LineCap.Round;
                p.EndCap = LineCap.Round;
                // Left arm
                g.DrawLine(p, sthCx-sz*0.10f, sthTop, sthCx-sz*0.03f, sthTop+sz*0.16f);
                // Right arm
                g.DrawLine(p, sthCx+sz*0.10f, sthTop, sthCx+sz*0.03f, sthTop+sz*0.16f);
                // Stem down
                g.DrawLine(p, sthCx, sthTop+sz*0.14f, sthCx, sthTop+sz*0.35f);
            }

            // Earpiece dots
            float dotR = Math.Max(sz*0.03f, 1.5f);
            using(var b = new SolidBrush(Color.FromArgb(200,200,200)))  {
                g.FillEllipse(b, sthCx-sz*0.10f-dotR, sthTop-dotR, dotR*2, dotR*2);
                g.FillEllipse(b, sthCx+sz*0.10f-dotR, sthTop-dotR, dotR*2, dotR*2);
            }

            // Tube (curved path from stem down to chest piece)
            using(var p = new Pen(Color.White, Math.Max(sz*0.035f, 1.5f))) {
                p.StartCap = LineCap.Round;
                p.EndCap = LineCap.Round;
                var tubePath = new GraphicsPath();
                float tubeStartY = sthTop+sz*0.35f;
                // Curve to the right and down to chest piece
                tubePath.AddBezier(
                    sthCx, tubeStartY,
                    sthCx+sz*0.22f, tubeStartY+sz*0.05f,
                    sthCx+sz*0.25f, tubeStartY+sz*0.25f,
                    sthCx+sz*0.08f, tubeStartY+sz*0.38f
                );
                g.DrawPath(p, tubePath);
            }

            // Chest piece (diaphragm)
            float cpX = sthCx+sz*0.08f, cpY = sthTop+sz*0.72f;
            float cpR = sz*0.07f;
            using(var b = new SolidBrush(Color.White))
                g.FillEllipse(b, cpX-cpR, cpY-cpR, cpR*2, cpR*2);
            using(var b = new SolidBrush(Color.FromArgb(150,34,139,230)))
                g.FillEllipse(b, cpX-cpR*0.6f, cpY-cpR*0.6f, cpR*1.2f, cpR*1.2f);
            // Highlight
            using(var b = new SolidBrush(Color.FromArgb(100,255,255,255)))
                g.FillEllipse(b, cpX-cpR*0.3f, cpY-cpR*0.5f, cpR*0.5f, cpR*0.4f);
        }
        return bmp;
    }

    static GraphicsPath RoundedRect(RectangleF r, float radius) {
        var p = new GraphicsPath();
        float d = radius*2;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right-d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right-d, r.Bottom-d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom-d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}
