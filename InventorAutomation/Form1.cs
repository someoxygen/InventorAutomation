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
                PartDocument partDoc = (PartDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                    _inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
                MessageBox.Show("Yeni parça oluþturuldu.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

        private void btnCreateSketch_Click(object sender, EventArgs e)
        {
            try
            {
                PartDocument partDoc = (PartDocument)_inventorApp.ActiveDocument;
                PartComponentDefinition compDef = partDoc.ComponentDefinition;
                PlanarSketch sketch = compDef.Sketches.Add(compDef.WorkPlanes[3]); // XY Plane

                TransientGeometry tg = _inventorApp.TransientGeometry;

                // Dikdörtgenin köþe noktalarý
                SketchPoint p1 = sketch.SketchPoints.Add(tg.CreatePoint2d(0, 0), false);
                SketchPoint p2 = sketch.SketchPoints.Add(tg.CreatePoint2d(10, 0), false);
                SketchPoint p3 = sketch.SketchPoints.Add(tg.CreatePoint2d(10, 5), false);
                SketchPoint p4 = sketch.SketchPoints.Add(tg.CreatePoint2d(0, 5), false);

                // Kenarlarý oluþtur
                sketch.SketchLines.AddByTwoPoints(p1, p2);
                sketch.SketchLines.AddByTwoPoints(p2, p3);
                sketch.SketchLines.AddByTwoPoints(p3, p4);
                sketch.SketchLines.AddByTwoPoints(p4, p1);

                MessageBox.Show("Sketch eklendi ve dikdörtgen çizildi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }
        private void btnSaveAndClose_Click(object sender, EventArgs e)
        {
            try
            {
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
