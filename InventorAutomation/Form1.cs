using System;
using System.Windows.Forms;
using Inventor;
namespace InventorAutomation
{
    public partial class Form1 : Form
    {

        private Inventor.Application _inventorApp;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnStartInventor_Click(object sender, EventArgs e)
        {
            try
            {
                btnCreatePart.Enabled = true;
                Type inventorType = Type.GetTypeFromProgID("Inventor.Application");
                _inventorApp = (Inventor.Application)Activator.CreateInstance(inventorType);
                _inventorApp.Visible = true;
                MessageBox.Show("Autodesk Inventor baþlatýldý.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

        private void btnCreatePart_Click(object sender, EventArgs e)
        {
            try
            {
                btnCreatePart.Enabled = false;
                btnRectangle.Enabled = true;
                btnTriangle.Enabled = true;
                btnCircle.Enabled = true;
                btnPentagon.Enabled = true;
                btnHexagon.Enabled = true;
                btnSaveAndClose.Enabled = true;
                PartDocument partDoc = (PartDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                    _inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
                MessageBox.Show("Yeni parça oluþturuldu.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }
        private void btnRectangle_Click(object sender, EventArgs e)
        {
            DrawShape(DrawRectangle, "Dikdörtgen çizildi.");
        }

        private void btnTriangle_Click(object sender, EventArgs e)
        {
            DrawShape(DrawTriangle, "Üçgen çizildi.");
        }

        private void btnCircle_Click(object sender, EventArgs e)
        {
            DrawShape(DrawCircle, "Daire çizildi.");
        }

        private void btnPentagon_Click(object sender, EventArgs e)
        {
            DrawShape(DrawPentagon, "Beþgen çizildi.");
        }

        private void btnHexagon_Click(object sender, EventArgs e)
        {
            DrawShape(DrawHexagon, "Altýgen çizildi.");
        }
        private void DrawShape(Action<PlanarSketch> drawMethod, string message)
        {
            try
            {
                PartDocument partDoc = (PartDocument)_inventorApp.ActiveDocument;
                PartComponentDefinition compDef = partDoc.ComponentDefinition;
                PlanarSketch sketch = compDef.Sketches.Add(compDef.WorkPlanes[3]); // XY Plane

                drawMethod(sketch);

                MessageBox.Show(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

        private void DrawRectangle(PlanarSketch sketch)
        {
            TransientGeometry tg = _inventorApp.TransientGeometry;

            SketchPoint p1 = sketch.SketchPoints.Add(tg.CreatePoint2d(0, 0), false);
            SketchPoint p2 = sketch.SketchPoints.Add(tg.CreatePoint2d(10, 0), false);
            SketchPoint p3 = sketch.SketchPoints.Add(tg.CreatePoint2d(10, 5), false);
            SketchPoint p4 = sketch.SketchPoints.Add(tg.CreatePoint2d(0, 5), false);

            sketch.SketchLines.AddByTwoPoints(p1, p2);
            sketch.SketchLines.AddByTwoPoints(p2, p3);
            sketch.SketchLines.AddByTwoPoints(p3, p4);
            sketch.SketchLines.AddByTwoPoints(p4, p1);
        }

        private void DrawTriangle(PlanarSketch sketch)
        {
            TransientGeometry tg = _inventorApp.TransientGeometry;

            SketchPoint p1 = sketch.SketchPoints.Add(tg.CreatePoint2d(0, 0), false);
            SketchPoint p2 = sketch.SketchPoints.Add(tg.CreatePoint2d(10, 0), false);
            SketchPoint p3 = sketch.SketchPoints.Add(tg.CreatePoint2d(5, 8), false);

            sketch.SketchLines.AddByTwoPoints(p1, p2);
            sketch.SketchLines.AddByTwoPoints(p2, p3);
            sketch.SketchLines.AddByTwoPoints(p3, p1);
        }

        private void DrawCircle(PlanarSketch sketch)
        {
            TransientGeometry tg = _inventorApp.TransientGeometry;
            sketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(5, 5), 5);
        }

        private void DrawPentagon(PlanarSketch sketch)
        {
            TransientGeometry tg = _inventorApp.TransientGeometry;
            double centerX = 5, centerY = 5, radius = 4;
            double angleStep = 2 * Math.PI / 5;

            SketchPoint[] points = new SketchPoint[5];

            for (int i = 0; i < 5; i++)
            {
                double x = centerX + radius * Math.Cos(i * angleStep);
                double y = centerY + radius * Math.Sin(i * angleStep);
                points[i] = sketch.SketchPoints.Add(tg.CreatePoint2d(x, y), false);
            }

            for (int i = 0; i < 5; i++)
            {
                sketch.SketchLines.AddByTwoPoints(points[i], points[(i + 1) % 5]);
            }
        }

        private void DrawHexagon(PlanarSketch sketch)
        {
            TransientGeometry tg = _inventorApp.TransientGeometry;
            double centerX = 5, centerY = 5, radius = 4;
            double angleStep = 2 * Math.PI / 6;

            SketchPoint[] points = new SketchPoint[6];

            for (int i = 0; i < 6; i++)
            {
                double x = centerX + radius * Math.Cos(i * angleStep);
                double y = centerY + radius * Math.Sin(i * angleStep);
                points[i] = sketch.SketchPoints.Add(tg.CreatePoint2d(x, y), false);
            }

            for (int i = 0; i < 6; i++)
            {
                sketch.SketchLines.AddByTwoPoints(points[i], points[(i + 1) % 6]);
            }
        }
        private void btnSaveAndClose_Click(object sender, EventArgs e)
        {
            try
            {
                btnCreatePart.Enabled = true;
                btnRectangle.Enabled = false;
                btnTriangle.Enabled = false;
                btnCircle.Enabled = false;
                btnPentagon.Enabled = false;
                btnHexagon.Enabled = false;
                btnSaveAndClose.Enabled = false;
                PartDocument partDoc = (PartDocument)_inventorApp.ActiveDocument;
                string directoryPath = "C:\\Temp\\";
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                string filePath = System.IO.Path.Combine(directoryPath, $"InventorPart_{DateTime.Now:yyyyMMdd_HHmmss}.ipt");
                partDoc.SaveAs(filePath, false);
                partDoc.Close(true);
                MessageBox.Show("Parça kaydedildi ve kapatýldý: " + filePath, "Baþarýlý", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
