using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using Path = System.IO.Path;
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
                btnCreateAllParts.Enabled = true;
                btnCreateWindowBig.Enabled = true;
                btnCreateWindowSmall.Enabled = true;
                btnAssembleWindowParts.Enabled = true;
                btnAssembleParts.Enabled = true;
                Type inventorType = Type.GetTypeFromProgID("Inventor.Application");
                _inventorApp = (Inventor.Application)Activator.CreateInstance(inventorType);
                _inventorApp.Visible = true;
                MessageBox.Show("Autodesk Inventor başlatıldı.");
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
                MessageBox.Show("Yeni parça oluşturuldu.");
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
            DrawShape(DrawPentagon, "Beşgen çizildi.");
        }

        private void btnHexagon_Click(object sender, EventArgs e)
        {
            DrawShape(DrawHexagon, "Altıgen çizildi.");
        }
        private void DrawShape(Action<PlanarSketch, PartDocument> drawMethod, string message)
        {
            try
            {
                PartDocument partDoc = (PartDocument)_inventorApp.ActiveDocument;
                PartComponentDefinition compDef = partDoc.ComponentDefinition;
                PlanarSketch sketch = compDef.Sketches.Add(compDef.WorkPlanes[3]); // XY Plane

                drawMethod(sketch, partDoc);

                MessageBox.Show(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }


        private void ExtrudeShape(PartDocument partDoc, PlanarSketch sketch)
        {
            // Extrude işlemi için profili al
            Profile profile = sketch.Profiles.AddForSolid();

            // Extrude parametrelerini ayarla (örneğin, 2 cm kalınlık)
            PartComponentDefinition compDef = partDoc.ComponentDefinition;
            ExtrudeDefinition extrudeDef = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(
                profile,
                PartFeatureOperationEnum.kJoinOperation // Katı oluşturma işlemi
            );

            // Kalınlığı belirle (1. yön, 2 cm)
            extrudeDef.SetDistanceExtent(
                2, // Kalınlık (cm ya da çizim birimine bağlı)
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection
            );

            // Extrude işlemini uygula
            ExtrudeFeature extrude = compDef.Features.ExtrudeFeatures.Add(extrudeDef);
        }

        private void DrawRectangle(PlanarSketch sketch, PartDocument partDoc)
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

            ExtrudeShape(partDoc, sketch);
        }

        private void DrawTriangle(PlanarSketch sketch, PartDocument partDoc)
        {
            TransientGeometry tg = _inventorApp.TransientGeometry;

            SketchPoint p1 = sketch.SketchPoints.Add(tg.CreatePoint2d(0, 0), false);
            SketchPoint p2 = sketch.SketchPoints.Add(tg.CreatePoint2d(10, 0), false);
            SketchPoint p3 = sketch.SketchPoints.Add(tg.CreatePoint2d(5, 8), false);

            sketch.SketchLines.AddByTwoPoints(p1, p2);
            sketch.SketchLines.AddByTwoPoints(p2, p3);
            sketch.SketchLines.AddByTwoPoints(p3, p1);

            ExtrudeShape(partDoc, sketch);
        }

        private void DrawCircle(PlanarSketch sketch, PartDocument partDoc)
        {
            TransientGeometry tg = _inventorApp.TransientGeometry;
            sketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(5, 5), 5);

            ExtrudeShape(partDoc, sketch);
        }

        private void DrawPentagon(PlanarSketch sketch, PartDocument partDoc)
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

            ExtrudeShape(partDoc, sketch);
        }

        private void DrawHexagon(PlanarSketch sketch, PartDocument partDoc)
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

            ExtrudeShape(partDoc, sketch);
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
                MessageBox.Show("Parça kaydedildi ve kapatıldı: " + filePath, "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAssembleParts_Click(object sender, EventArgs e)
        {
            try
            {
                string directoryPath = "C:\\Temp\\";
                if (!Directory.Exists(directoryPath))
                    Directory.CreateDirectory(directoryPath);

                string partPath1 = Path.Combine(directoryPath, "Part1.ipt");
                string partPath2 = Path.Combine(directoryPath, "Part2.ipt");
                string partPath3 = Path.Combine(directoryPath, "Part3.ipt");
                string partPath4 = Path.Combine(directoryPath, "Part4.ipt");

                CreateBasicRectanglePart(partPath1, 100, 10);
                CreateBasicRectanglePart(partPath3, 100, 10);

                CreateBasicRectanglePart(partPath2, 80, 10);
                CreateBasicRectanglePart(partPath4, 80, 10);

                AssemblyDocument asmDoc = (AssemblyDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, "", true);
                AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
                TransientGeometry tg = _inventorApp.TransientGeometry;

                Matrix m1 = tg.CreateMatrix();
                ComponentOccurrence occ1 = asmDef.Occurrences.Add(partPath1, m1);
                occ1.Grounded = true;

                Matrix m2 = tg.CreateMatrix();
                m2.SetToRotation(Math.PI / 2, tg.CreateVector(0, 0, 1), tg.CreatePoint(0, 0, 0));
                m2.Cell[1, 4] = 100;
                ComponentOccurrence occ2 = asmDef.Occurrences.Add(partPath2, m2);

                Matrix m3 = tg.CreateMatrix();
                m3.Cell[2, 4] = -100;
                ComponentOccurrence occ3 = asmDef.Occurrences.Add(partPath3, m3);

                Matrix m4 = tg.CreateMatrix();
                m4.SetToRotation(Math.PI / 2, tg.CreateVector(0, 0, 1), tg.CreatePoint(0, 0, 0));
                m4.Cell[1, 4] = 100;
                m4.Cell[2, 4] = -100;
                ComponentOccurrence occ4 = asmDef.Occurrences.Add(partPath4, m4);

                AddMate(asmDef, occ1, "MateTop", occ2, "MateLeft",0);
                
                WorkPoint asmCenterPoint = asmDef.WorkPoints["Center Point"];

                object plane2Proxy;
                occ2.CreateGeometryProxy(((PartComponentDefinition)occ2.Definition).WorkPlanes["MateTop"], out plane2Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane2Proxy, 0);

                AddMate(asmDef, occ1, "MateTop", occ4, "MateLeft",0);

                object plane4Proxy;
                occ4.CreateGeometryProxy(((PartComponentDefinition)occ4.Definition).WorkPlanes["MateTop"], out plane4Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane4Proxy, 90);

                AddMate(asmDef, occ3, "MateBottom", occ2, "MateRight",0);

                object plane3Proxy;
                occ3.CreateGeometryProxy(((PartComponentDefinition)occ3.Definition).WorkPlanes["MateLeft"], out plane3Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane3Proxy, 0);

                string asmPath = Path.Combine(directoryPath, $"KareMontaj_{DateTime.Now:yyyyMMdd_HHmmss}.iam");
                asmDoc.SaveAs(asmPath, false);
                MessageBox.Show("Kare başarıyla oluşturuldu:\n" + asmPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }



        private void AddMate(AssemblyComponentDefinition asmDef, ComponentOccurrence occA, string wpA, ComponentOccurrence occB, string wpB, int offset)
        {
            object proxyA, proxyB;
            occA.CreateGeometryProxy(((PartComponentDefinition)occA.Definition).WorkPlanes[wpA], out proxyA);
            occB.CreateGeometryProxy(((PartComponentDefinition)occB.Definition).WorkPlanes[wpB], out proxyB);
            asmDef.Constraints.AddMateConstraint((WorkPlaneProxy)proxyA, (WorkPlaneProxy)proxyB, offset);
        }

        private void AddFlush(AssemblyComponentDefinition asmDef, ComponentOccurrence occA, string wpA, ComponentOccurrence occB, string wpB, int offset)
        {
            object proxyA, proxyB;
            occA.CreateGeometryProxy(((PartComponentDefinition)occA.Definition).WorkPlanes[wpA], out proxyA);
            occB.CreateGeometryProxy(((PartComponentDefinition)occB.Definition).WorkPlanes[wpB], out proxyB);
            asmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)proxyA, (WorkPlaneProxy)proxyB, offset);
        }


        private void CreateBasicRectanglePart(string path, double length, double height)
        {
            PartDocument partDoc = (PartDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, "", true);
            PartComponentDefinition partDef = partDoc.ComponentDefinition;
            TransientGeometry tg = _inventorApp.TransientGeometry;

            // Dikdörtgen çiz
            PlanarSketch sketch = partDef.Sketches.Add(partDef.WorkPlanes[3]); // XY Plane
            sketch.SketchLines.AddAsTwoPointRectangle(tg.CreatePoint2d(0, 0), tg.CreatePoint2d(length, height));

            // Extrude (katı oluştur)
            Profile profile = sketch.Profiles.AddForSolid();
            ExtrudeFeature mainExtrude = partDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                profile,
                10, // derinlik
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection,
                PartFeatureOperationEnum.kJoinOperation,
                null
            );

            // ✅ Delik için sketch, bu sefer extrude edilmiş üst yüzeye açılıyor (Yüzeye değil, Plane'e sketch yaparsan görünmeyebilir)
            Face topFace = partDef.SurfaceBodies[1].Faces
                .Cast<Face>()
                .FirstOrDefault(f =>
                    f.SurfaceType == SurfaceTypeEnum.kPlaneSurface &&
                    Math.Abs(((Plane)f.Geometry).Normal.Z - 1) < 0.01); // Z-ekseni yönünde yüzeyi bul (üst yüz)

            if (topFace != null)
            {
                PlanarSketch holeSketch = partDef.Sketches.Add(topFace);

                double holeRadius = 1.5;
                double edgeOffset = 5.0;

                // 4 köşeye daire delikleri
                holeSketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(edgeOffset, edgeOffset), holeRadius);
                holeSketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(length - edgeOffset, edgeOffset), holeRadius);
                holeSketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(edgeOffset, height - edgeOffset), holeRadius);
                holeSketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(length - edgeOffset, height - edgeOffset), holeRadius);

                Profile holeProfile = holeSketch.Profiles.AddForSolid();

                partDef.Features.ExtrudeFeatures.AddByDistanceExtent(
                    holeProfile,
                    10, // delik derinliği (tam katı boyunca gitsin)
                    PartFeatureExtentDirectionEnum.kNegativeExtentDirection,
                    PartFeatureOperationEnum.kCutOperation,
                    null
                );
            }

            // -- WorkPlane'ler (Mate işlemleri için) --
            void AddWorkPlaneByFace(Func<Plane, bool> normalCheck, string name)
            {
                var face = partDef.SurfaceBodies[1].Faces
                    .Cast<Face>()
                    .FirstOrDefault(f => f.SurfaceType == SurfaceTypeEnum.kPlaneSurface && f.Geometry is Plane plane && normalCheck(plane));
                if (face != null)
                {
                    var wp = partDef.WorkPlanes.AddByPlaneAndOffset(face, 0);
                    wp.Name = name;
                }
            }

            AddWorkPlaneByFace(p => Math.Abs(p.Normal.X - 1) < 0.01, "MateRight");
            AddWorkPlaneByFace(p => Math.Abs(p.Normal.X + 1) < 0.01, "MateLeft");
            AddWorkPlaneByFace(p => Math.Abs(p.Normal.Y - 1) < 0.01, "MateTop");
            AddWorkPlaneByFace(p => Math.Abs(p.Normal.Y + 1) < 0.01, "MateBottom");

            partDoc.SaveAs(path, false);
            partDoc.Close();
        }



        private void btnAssembleWindow_Click(object sender, EventArgs e)
        {
            try
            {
                string directoryPath = "C:\\Temp\\";
                if (!Directory.Exists(directoryPath))
                    Directory.CreateDirectory(directoryPath);

                //CREATE PARTS REGION

                string partPath1 = Path.Combine(directoryPath, "Part1.ipt");
                string partPath2 = Path.Combine(directoryPath, "Part2.ipt");
                string partPath3 = Path.Combine(directoryPath, "Part3.ipt");
                string partPath4 = Path.Combine(directoryPath, "Part4.ipt");
                string partPath5 = Path.Combine(directoryPath, "Part5.ipt");
                string partPath6 = Path.Combine(directoryPath, "Part6.ipt");
                string partPath7 = Path.Combine(directoryPath, "Part7.ipt");
                string partPath8 = Path.Combine(directoryPath, "Part8.ipt");
                string partPath9 = Path.Combine(directoryPath, "Part9.ipt");
                string partPath10 = Path.Combine(directoryPath, "Part10.ipt");
                string partPath11 = Path.Combine(directoryPath, "Part11.ipt");
                string partPath12 = Path.Combine(directoryPath, "Part12.ipt");

                CreateBasicRectanglePart(partPath1, 100, 10);
                CreateBasicRectanglePart(partPath3, 100, 10);
                CreateBasicRectanglePart(partPath5, 100, 10);
                CreateBasicRectanglePart(partPath7, 100, 10);
                CreateBasicRectanglePart(partPath9, 100, 10);
                CreateBasicRectanglePart(partPath11, 100, 10);

                CreateBasicRectanglePart(partPath2, 80, 10);
                CreateBasicRectanglePart(partPath4, 80, 10);
                CreateBasicRectanglePart(partPath6, 80, 10);
                CreateBasicRectanglePart(partPath8, 80, 10);
                CreateBasicRectanglePart(partPath10, 80, 10);
                CreateBasicRectanglePart(partPath12, 80, 10);

                AssemblyDocument asmDoc = (AssemblyDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, "", true);
                AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;
                TransientGeometry tg = _inventorApp.TransientGeometry;


                //MATRIX REGION

                // İlk parça (0,0,0) - yatay
                Matrix m1 = tg.CreateMatrix();
                ComponentOccurrence occ1 = asmDef.Occurrences.Add(partPath1, m1);
                occ1.Grounded = true;

                // İkinci parça (0,100,0) - dikey
                Matrix m2 = tg.CreateMatrix();
                m2.SetToRotation(Math.PI / 2, tg.CreateVector(0, 0, 1), tg.CreatePoint(0, 0, 0));
                m2.Cell[1, 4] = 100;
                ComponentOccurrence occ2 = asmDef.Occurrences.Add(partPath2, m2);

                // Üçüncü parça (0,0,-100) - yatay
                Matrix m3 = tg.CreateMatrix();
                m3.Cell[2, 4] = -100;
                ComponentOccurrence occ3 = asmDef.Occurrences.Add(partPath3, m3);

                // Dördüncü parça (0,100,-100) - dikey
                Matrix m4 = tg.CreateMatrix();
                m4.SetToRotation(Math.PI / 2, tg.CreateVector(0, 0, 1), tg.CreatePoint(0, 0, 0));
                m4.Cell[1, 4] = 100;
                m4.Cell[2, 4] = -100;
                ComponentOccurrence occ4 = asmDef.Occurrences.Add(partPath4, m4);

                // Beşinci parça (0,0,-200) - yatay
                Matrix m5 = tg.CreateMatrix();
                m5.Cell[2, 4] = -200;
                ComponentOccurrence occ5 = asmDef.Occurrences.Add(partPath5, m5);

                // Altıncı parça (0,100,-200) - dikey
                Matrix m6 = tg.CreateMatrix();
                m6.SetToRotation(Math.PI / 2, tg.CreateVector(0, 0, 1), tg.CreatePoint(0, 0, 0));
                m6.Cell[1, 4] = 100;
                m6.Cell[2, 4] = -200;
                ComponentOccurrence occ6 = asmDef.Occurrences.Add(partPath6, m6);

                // Yedinci parça (0,0,-300) - yatay
                Matrix m7 = tg.CreateMatrix();
                m7.Cell[2, 4] = -300;
                ComponentOccurrence occ7 = asmDef.Occurrences.Add(partPath7, m7);

                // Sekizinci parça (0,100,-300) - dikey
                Matrix m8 = tg.CreateMatrix();
                m8.SetToRotation(Math.PI / 2, tg.CreateVector(0, 0, 1), tg.CreatePoint(0, 0, 0));
                m8.Cell[1, 4] = 100;
                m8.Cell[2, 4] = -300;
                ComponentOccurrence occ8 = asmDef.Occurrences.Add(partPath8, m8);

                // Dokuzuncu parça (0,0,-400) - yatay
                Matrix m9 = tg.CreateMatrix();
                m9.Cell[2, 4] = -400;
                ComponentOccurrence occ9 = asmDef.Occurrences.Add(partPath9, m9);

                // Onuncu parça (0,100,-400) - dikey
                Matrix m10 = tg.CreateMatrix();
                m10.SetToRotation(Math.PI / 2, tg.CreateVector(0, 0, 1), tg.CreatePoint(0, 0, 0));
                m10.Cell[1, 4] = 100;
                m10.Cell[2, 4] = -400;
                ComponentOccurrence occ10 = asmDef.Occurrences.Add(partPath10, m10);

                // On birinci parça (0,0,-500) - yatay
                Matrix m11 = tg.CreateMatrix();
                m11.Cell[2, 4] = -500;
                ComponentOccurrence occ11 = asmDef.Occurrences.Add(partPath11, m11);

                // On ikinci parça (0,100,-500) - dikey
                Matrix m12 = tg.CreateMatrix();
                m12.SetToRotation(Math.PI / 2, tg.CreateVector(0, 0, 1), tg.CreatePoint(0, 0, 0));
                m12.Cell[1, 4] = 100;
                m12.Cell[2, 4] = -500;
                ComponentOccurrence occ12 = asmDef.Occurrences.Add(partPath12, m12);


                // CONSTRAINT REGION

                AddMate(asmDef, occ1, "MateTop", occ2, "MateLeft", 0);

                WorkPoint asmCenterPoint = asmDef.WorkPoints["Center Point"];

                object plane2Proxy;
                occ2.CreateGeometryProxy(((PartComponentDefinition)occ2.Definition).WorkPlanes["MateTop"], out plane2Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane2Proxy, 0);

                AddMate(asmDef, occ1, "MateTop", occ4, "MateLeft", 0);

                object plane4Proxy;
                occ4.CreateGeometryProxy(((PartComponentDefinition)occ4.Definition).WorkPlanes["MateTop"], out plane4Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane4Proxy, 90);

                AddMate(asmDef, occ3, "MateBottom", occ2, "MateRight", 0);

                object plane3Proxy;
                occ3.CreateGeometryProxy(((PartComponentDefinition)occ3.Definition).WorkPlanes["MateLeft"], out plane3Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane3Proxy, 0);

                AddMate(asmDef, occ5, "MateLeft", occ1, "MateRight", 0);

                AddMate(asmDef, occ4, "MateBottom", occ6, "MateTop", 90);

                AddMate(asmDef, occ7, "MateLeft", occ3, "MateRight", 0);

                AddMate(asmDef, occ12, "MateLeft", occ3, "MateTop", 0);

                object plane12Proxy;
                occ12.CreateGeometryProxy(((PartComponentDefinition)occ12.Definition).WorkPlanes["MateTop"], out plane12Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane12Proxy, 0);

                AddMate(asmDef, occ12, "MateBottom", occ10, "MateTop", 80);

                object plane10Proxy;
                occ10.CreateGeometryProxy(((PartComponentDefinition)occ10.Definition).WorkPlanes["MateLeft"], out plane10Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane10Proxy, 100);

                AddMate(asmDef, occ10, "MateTop", occ8, "MateBottom", -110);

                AddMate(asmDef, occ12, "MateRight", occ11, "MateBottom", 0);

                object plane11Proxy;
                occ11.CreateGeometryProxy(((PartComponentDefinition)occ11.Definition).WorkPlanes["MateLeft"], out plane11Proxy);

                asmDef.Constraints.AddMateConstraint(asmCenterPoint, (WorkPlaneProxy)plane11Proxy, 0);

                AddMate(asmDef, occ9, "MateLeft", occ11, "MateRight", 0);

                string asmPath = Path.Combine(directoryPath, $"KareMontaj_{DateTime.Now:yyyyMMdd_HHmmss}.iam");
                asmDoc.SaveAs(asmPath, false);
                MessageBox.Show("Kare başarıyla oluşturuldu:\n" + asmPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message);
            }
        }

        private void btnCreateBigWindow_Click(object sender, EventArgs e)
        {
            try
            {
                string assemblyToInsertPath = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol1\60ST15_Kasa.iam";

                if (!System.IO.File.Exists(assemblyToInsertPath))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPath, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 1. Boş bir montaj oluştur
                AssemblyDocument newAsmDoc = (AssemblyDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, "", true);
                AssemblyComponentDefinition newAsmDef = newAsmDoc.ComponentDefinition;

                // 2. Dosyayı montaj ortamına ekle
                Matrix insertMatrix = _inventorApp.TransientGeometry.CreateMatrix();
                ComponentOccurrence insertedOcc1 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);

                // 3. Profilin XY Plane'ini montajın XY Plane'iyle hizala (Flush)
                WorkPlane asmXYPlane = newAsmDef.WorkPlanes["XY Plane"];
                WorkPlane occ1XYPlane = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPlanes["XY Plane"];
                object occ1XYPlaneProxyObj;
                insertedOcc1.CreateGeometryProxy(occ1XYPlane, out occ1XYPlaneProxyObj);
                insertedOcc1.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ1XYPlaneProxyObj, asmXYPlane, 0);
                
                // 4. Kullanıcıdan offset değerleri (örnek değerler burada sabit)
                double offsetXOcc1 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc1 = GetParameterValue(insertedOcc1, "Length"); // aşağıdan yükseklik (mm)

                // 5. WorkPoint1'i al ve proxy'sini oluştur
                WorkPoint workPtOcc1 = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc1;
                insertedOcc1.CreateGeometryProxy(workPtOcc1, out workPointProxyObjOcc1);
                WorkPointProxy workPointProxyOcc1 = (WorkPointProxy)workPointProxyObjOcc1;
                
                // 6. Mate ile XZ Plane (Yükseklik)
                WorkPlane asmXZPlane = newAsmDef.WorkPlanes["XZ Plane"];
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc1, asmXZPlane, offsetYOcc1);
                
                // 7. Mate ile YZ Plane (Soldan uzaklık)
                WorkPlane asmYZPlane = newAsmDef.WorkPlanes["YZ Plane"];
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc1, asmYZPlane, offsetXOcc1);
                
                // 8. Work Plane5'i al ve proxy'sini oluştur
                WorkPlane occWorkPlane5Occ1 = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc1;
                insertedOcc1.CreateGeometryProxy(occWorkPlane5Occ1, out workPlane5ProxyObjOcc1);
                WorkPlaneProxy workPlane5ProxyOcc1 = (WorkPlaneProxy)workPlane5ProxyObjOcc1;
                
                // 9. DirectedAngleConstraint ile 0 derece açı tanımla (yatay montaj)
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc1,asmXZPlane, 0);
                

                //OCC2
                ComponentOccurrence insertedOcc2 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ2XYPlane = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPlanes["XY Plane"];
                object occ2XYPlaneProxyObj;
                insertedOcc2.CreateGeometryProxy(occ2XYPlane, out occ2XYPlaneProxyObj);
                insertedOcc2.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ2XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc2 = 50.0 - GetParameterValue(insertedOcc2, "Length") / 2; // soldan uzaklık (mm)
                double offsetYOcc2 = GetParameterValue(insertedOcc2, "Length") + GetParameterValue(insertedOcc2, "Length") / 2; // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc2 = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc2;
                insertedOcc2.CreateGeometryProxy(workPtOcc2, out workPointProxyObjOcc2);
                WorkPointProxy workPointProxyOcc2 = (WorkPointProxy)workPointProxyObjOcc2;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc2, asmXZPlane, offsetYOcc2);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc2, asmYZPlane, offsetXOcc2);
                WorkPlane occWorkPlane5Occ2 = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc2;
                insertedOcc2.CreateGeometryProxy(occWorkPlane5Occ2, out workPlane5ProxyObjOcc2);
                WorkPlaneProxy workPlane5ProxyOcc2 = (WorkPlaneProxy)workPlane5ProxyObjOcc2;
                double angleOcc2InRadians = 90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc2, asmXZPlane, angleOcc2InRadians);


                //OCC3
                ComponentOccurrence insertedOcc3 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ3XYPlane = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPlanes["XY Plane"];
                object occ3XYPlaneProxyObj;
                insertedOcc3.CreateGeometryProxy(occ3XYPlane, out occ3XYPlaneProxyObj);
                insertedOcc3.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ3XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc3 = 50.0 + GetParameterValue(insertedOcc3, "Length") / 2; // soldan uzaklık (mm)
                double offsetYOcc3 = GetParameterValue(insertedOcc3, "Length") + GetParameterValue(insertedOcc3, "Length") / 2; // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc3 = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc3;
                insertedOcc3.CreateGeometryProxy(workPtOcc3, out workPointProxyObjOcc3);
                WorkPointProxy workPointProxyOcc3 = (WorkPointProxy)workPointProxyObjOcc3;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc3, asmXZPlane, offsetYOcc3);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc3, asmYZPlane, offsetXOcc3);
                WorkPlane occWorkPlane5Occ3 = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc3;
                insertedOcc3.CreateGeometryProxy(occWorkPlane5Occ3, out workPlane5ProxyObjOcc3);
                WorkPlaneProxy workPlane5ProxyOcc3 = (WorkPlaneProxy)workPlane5ProxyObjOcc3;
                double angleOcc3InRadians = -90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc3, asmXZPlane, angleOcc3InRadians);


                //OCC4
                ComponentOccurrence insertedOcc4 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ4XYPlane = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPlanes["XY Plane"];
                object occ4XYPlaneProxyObj;
                insertedOcc4.CreateGeometryProxy(occ4XYPlane, out occ4XYPlaneProxyObj);
                insertedOcc4.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ4XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc4 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc4 = GetParameterValue(insertedOcc4, "Length") + GetParameterValue(insertedOcc4, "Length"); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc4 = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc4;
                insertedOcc4.CreateGeometryProxy(workPtOcc4, out workPointProxyObjOcc4);
                WorkPointProxy workPointProxyOcc4 = (WorkPointProxy)workPointProxyObjOcc4;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc4, asmXZPlane, offsetYOcc4);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc4, asmYZPlane, offsetXOcc4);
                WorkPlane occWorkPlane5Occ4 = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc4;
                insertedOcc4.CreateGeometryProxy(occWorkPlane5Occ4, out workPlane5ProxyObjOcc4);
                WorkPlaneProxy workPlane5ProxyOcc4 = (WorkPlaneProxy)workPlane5ProxyObjOcc4;
                double angleOcc4InRadians = 180 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc4, asmXZPlane, angleOcc4InRadians);

                //OCC5
                ComponentOccurrence insertedOcc5 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ5XYPlane = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPlanes["XY Plane"];
                object occ5XYPlaneProxyObj;
                insertedOcc5.CreateGeometryProxy(occ5XYPlane, out occ5XYPlaneProxyObj);
                insertedOcc5.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ5XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc5 = 50.0 + GetParameterValue(insertedOcc5, "Length"); // soldan uzaklık (mm)
                double offsetYOcc5 = GetParameterValue(insertedOcc5, "Length") ; // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc5 = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc5;
                insertedOcc5.CreateGeometryProxy(workPtOcc5, out workPointProxyObjOcc5);
                WorkPointProxy workPointProxyOcc5 = (WorkPointProxy)workPointProxyObjOcc5;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc5, asmXZPlane, offsetYOcc5);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc5, asmYZPlane, offsetXOcc5);
                WorkPlane occWorkPlane5Occ5 = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc5;
                insertedOcc5.CreateGeometryProxy(occWorkPlane5Occ5, out workPlane5ProxyObjOcc5);
                WorkPlaneProxy workPlane5ProxyOcc5 = (WorkPlaneProxy)workPlane5ProxyObjOcc5;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc5, asmXZPlane, 0);

                //OCC6
                ComponentOccurrence insertedOcc6 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ6XYPlane = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPlanes["XY Plane"];
                object occ6XYPlaneProxyObj;
                insertedOcc6.CreateGeometryProxy(occ6XYPlane, out occ6XYPlaneProxyObj);
                insertedOcc6.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ6XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc6 = 50.0 + GetParameterValue(insertedOcc6, "Length") + (GetParameterValue(insertedOcc6, "Length") / 2); // soldan uzaklık (mm)
                double offsetYOcc6 = GetParameterValue(insertedOcc6, "Length") + (GetParameterValue(insertedOcc6, "Length") / 2); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc6 = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc6;
                insertedOcc6.CreateGeometryProxy(workPtOcc6, out workPointProxyObjOcc6);
                WorkPointProxy workPointProxyOcc6 = (WorkPointProxy)workPointProxyObjOcc6;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc6, asmXZPlane, offsetYOcc6);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc6, asmYZPlane, offsetXOcc6);
                WorkPlane occWorkPlane5Occ6 = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc6;
                insertedOcc6.CreateGeometryProxy(occWorkPlane5Occ6, out workPlane5ProxyObjOcc6);
                WorkPlaneProxy workPlane5ProxyOcc6 = (WorkPlaneProxy)workPlane5ProxyObjOcc6;
                double angleOcc6InRadians = -90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc6, asmXZPlane, angleOcc6InRadians);


                //OCC7
                ComponentOccurrence insertedOcc7 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ7XYPlane = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPlanes["XY Plane"];
                object occ7XYPlaneProxyObj;
                insertedOcc7.CreateGeometryProxy(occ7XYPlane, out occ7XYPlaneProxyObj);
                insertedOcc7.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ7XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc7 = 50.0 + GetParameterValue(insertedOcc7, "Length"); // soldan uzaklık (mm)
                double offsetYOcc7 = GetParameterValue(insertedOcc7, "Length") + GetParameterValue(insertedOcc7, "Length"); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc7 = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc7;
                insertedOcc7.CreateGeometryProxy(workPtOcc7, out workPointProxyObjOcc7);
                WorkPointProxy workPointProxyOcc7 = (WorkPointProxy)workPointProxyObjOcc7;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc7, asmXZPlane, offsetYOcc7);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc7, asmYZPlane, offsetXOcc7);
                WorkPlane occWorkPlane5Occ7 = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc7;
                insertedOcc7.CreateGeometryProxy(occWorkPlane5Occ7, out workPlane5ProxyObjOcc7);
                WorkPlaneProxy workPlane5ProxyOcc7 = (WorkPlaneProxy)workPlane5ProxyObjOcc7;
                double angleOcc7InRadians = -180 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc7, asmXZPlane, angleOcc7InRadians);

                //OCC8
                ComponentOccurrence insertedOcc8 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ8XYPlane = ((AssemblyComponentDefinition)insertedOcc8.Definition).WorkPlanes["XY Plane"];
                object occ8XYPlaneProxyObj;
                insertedOcc8.CreateGeometryProxy(occ8XYPlane, out occ8XYPlaneProxyObj);
                insertedOcc8.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ8XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc8 = 50.0 + GetParameterValue(insertedOcc8, "Length") + (GetParameterValue(insertedOcc8, "Length") / 2); // soldan uzaklık (mm)
                double offsetYOcc8 = GetParameterValue(insertedOcc8, "Length") + GetParameterValue(insertedOcc8, "Length") + (GetParameterValue(insertedOcc8, "Length") /2); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc8 = ((AssemblyComponentDefinition)insertedOcc8.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc8;
                insertedOcc8.CreateGeometryProxy(workPtOcc8, out workPointProxyObjOcc8);
                WorkPointProxy workPointProxyOcc8 = (WorkPointProxy)workPointProxyObjOcc8;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc8, asmXZPlane, offsetYOcc8);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc8, asmYZPlane, offsetXOcc8);
                WorkPlane occWorkPlane5Occ8 = ((AssemblyComponentDefinition)insertedOcc8.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc8;
                insertedOcc8.CreateGeometryProxy(occWorkPlane5Occ8, out workPlane5ProxyObjOcc8);
                WorkPlaneProxy workPlane5ProxyOcc8 = (WorkPlaneProxy)workPlane5ProxyObjOcc8;
                double angleOcc8InRadians = -90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc8, asmXZPlane, angleOcc8InRadians);

                //OCC9
                ComponentOccurrence insertedOcc9 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ9XYPlane = ((AssemblyComponentDefinition)insertedOcc9.Definition).WorkPlanes["XY Plane"];
                object occ9XYPlaneProxyObj;
                insertedOcc9.CreateGeometryProxy(occ9XYPlane, out occ9XYPlaneProxyObj);
                insertedOcc9.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ9XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc9 = 50.0 + GetParameterValue(insertedOcc9, "Length"); // soldan uzaklık (mm)
                double offsetYOcc9 = GetParameterValue(insertedOcc9, "Length") + GetParameterValue(insertedOcc9, "Length") + GetParameterValue(insertedOcc9, "Length"); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc9 = ((AssemblyComponentDefinition)insertedOcc9.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc9;
                insertedOcc9.CreateGeometryProxy(workPtOcc9, out workPointProxyObjOcc9);
                WorkPointProxy workPointProxyOcc9 = (WorkPointProxy)workPointProxyObjOcc9;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc9, asmXZPlane, offsetYOcc9);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc9, asmYZPlane, offsetXOcc9);
                WorkPlane occWorkPlane5Occ9 = ((AssemblyComponentDefinition)insertedOcc9.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc9;
                insertedOcc9.CreateGeometryProxy(occWorkPlane5Occ9, out workPlane5ProxyObjOcc9);
                WorkPlaneProxy workPlane5ProxyOcc9 = (WorkPlaneProxy)workPlane5ProxyObjOcc9;
                double angleOcc9InRadians = -180 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc9, asmXZPlane, angleOcc9InRadians);

                //OCC10
                ComponentOccurrence insertedOcc10 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ10XYPlane = ((AssemblyComponentDefinition)insertedOcc10.Definition).WorkPlanes["XY Plane"];
                object occ10XYPlaneProxyObj;
                insertedOcc10.CreateGeometryProxy(occ10XYPlane, out occ10XYPlaneProxyObj);
                insertedOcc10.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ10XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc10 = 50.0 + GetParameterValue(insertedOcc10, "Length") / 2; // soldan uzaklık (mm)
                double offsetYOcc10 = GetParameterValue(insertedOcc10, "Length") + GetParameterValue(insertedOcc10, "Length") + GetParameterValue(insertedOcc10, "Length") / 2; // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc10 = ((AssemblyComponentDefinition)insertedOcc10.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc10;
                insertedOcc10.CreateGeometryProxy(workPtOcc10, out workPointProxyObjOcc10);
                WorkPointProxy workPointProxyOcc10 = (WorkPointProxy)workPointProxyObjOcc10;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc10, asmXZPlane, offsetYOcc10);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc10, asmYZPlane, offsetXOcc10);
                WorkPlane occWorkPlane5Occ10 = ((AssemblyComponentDefinition)insertedOcc10.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc10;
                insertedOcc10.CreateGeometryProxy(occWorkPlane5Occ10, out workPlane5ProxyObjOcc10);
                WorkPlaneProxy workPlane5ProxyOcc10 = (WorkPlaneProxy)workPlane5ProxyObjOcc10;
                double angleOcc10InRadians = -90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc10, asmXZPlane, angleOcc10InRadians);

                //OCC11
                ComponentOccurrence insertedOcc11 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ11XYPlane = ((AssemblyComponentDefinition)insertedOcc11.Definition).WorkPlanes["XY Plane"];
                object occ11XYPlaneProxyObj;
                insertedOcc11.CreateGeometryProxy(occ11XYPlane, out occ11XYPlaneProxyObj);
                insertedOcc11.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ11XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc11 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc11 = GetParameterValue(insertedOcc11, "Length") + GetParameterValue(insertedOcc11, "Length") + GetParameterValue(insertedOcc11, "Length"); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc11 = ((AssemblyComponentDefinition)insertedOcc11.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc11;
                insertedOcc11.CreateGeometryProxy(workPtOcc11, out workPointProxyObjOcc11);
                WorkPointProxy workPointProxyOcc11 = (WorkPointProxy)workPointProxyObjOcc11;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc11, asmXZPlane, offsetYOcc11);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc11, asmYZPlane, offsetXOcc11);
                WorkPlane occWorkPlane5Occ11 = ((AssemblyComponentDefinition)insertedOcc11.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc11;
                insertedOcc11.CreateGeometryProxy(occWorkPlane5Occ11, out workPlane5ProxyObjOcc11);
                WorkPlaneProxy workPlane5ProxyOcc11 = (WorkPlaneProxy)workPlane5ProxyObjOcc11;
                double angleOcc11InRadians = 180 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc11, asmXZPlane, angleOcc11InRadians);

                //OCC12
                ComponentOccurrence insertedOcc12 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ12XYPlane = ((AssemblyComponentDefinition)insertedOcc12.Definition).WorkPlanes["XY Plane"];
                object occ12XYPlaneProxyObj;
                insertedOcc12.CreateGeometryProxy(occ12XYPlane, out occ12XYPlaneProxyObj);
                insertedOcc12.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ12XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc12 = 50.0 - GetParameterValue(insertedOcc12, "Length") / 2; // soldan uzaklık (mm)
                double offsetYOcc12 = GetParameterValue(insertedOcc12, "Length") + GetParameterValue(insertedOcc12, "Length") + (GetParameterValue(insertedOcc12, "Length") / 2); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc12 = ((AssemblyComponentDefinition)insertedOcc12.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc12;
                insertedOcc12.CreateGeometryProxy(workPtOcc12, out workPointProxyObjOcc12);
                WorkPointProxy workPointProxyOcc12 = (WorkPointProxy)workPointProxyObjOcc12;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc12, asmXZPlane, offsetYOcc12);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc12, asmYZPlane, offsetXOcc12);
                WorkPlane occWorkPlane5Occ12 = ((AssemblyComponentDefinition)insertedOcc12.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc12;
                insertedOcc12.CreateGeometryProxy(occWorkPlane5Occ12, out workPlane5ProxyObjOcc12);
                WorkPlaneProxy workPlane5ProxyOcc12 = (WorkPlaneProxy)workPlane5ProxyObjOcc12;
                double angleOcc12InRadians = 90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc12, asmXZPlane, angleOcc12InRadians);

                MessageBox.Show("Montaj başarıyla eklendi, hizalandı ve konumlandırıldı!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Yeni montaj dosyasına eklenirken hata oluştu:\n" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCreateSmallWindow_Click(object sender, EventArgs e)
        {
            try
            {
                string assemblyToInsertPath = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol1\60ST15_Kasa.iam";

                if (!System.IO.File.Exists(assemblyToInsertPath))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPath, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 1. Boş bir montaj oluştur
                AssemblyDocument newAsmDoc = (AssemblyDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, "", true);
                AssemblyComponentDefinition newAsmDef = newAsmDoc.ComponentDefinition;

                // 2. Dosyayı montaj ortamına ekle
                Matrix insertMatrix = _inventorApp.TransientGeometry.CreateMatrix();
                ComponentOccurrence insertedOcc1 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);

                // 3. Profilin XY Plane'ini montajın XY Plane'iyle hizala (Flush)
                WorkPlane asmXYPlane = newAsmDef.WorkPlanes["XY Plane"];
                WorkPlane occ1XYPlane = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPlanes["XY Plane"];
                object occ1XYPlaneProxyObj;
                insertedOcc1.CreateGeometryProxy(occ1XYPlane, out occ1XYPlaneProxyObj);
                insertedOcc1.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ1XYPlaneProxyObj, asmXYPlane, 0);

                // 4. Kullanıcıdan offset değerleri (örnek değerler burada sabit)
                double offsetXOcc1 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc1 = GetParameterValue(insertedOcc1, "Length"); // aşağıdan yükseklik (mm)

                // 5. WorkPoint1'i al ve proxy'sini oluştur
                WorkPoint workPtOcc1 = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc1;
                insertedOcc1.CreateGeometryProxy(workPtOcc1, out workPointProxyObjOcc1);
                WorkPointProxy workPointProxyOcc1 = (WorkPointProxy)workPointProxyObjOcc1;

                // 6. Mate ile XZ Plane (Yükseklik)
                WorkPlane asmXZPlane = newAsmDef.WorkPlanes["XZ Plane"];
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc1, asmXZPlane, offsetYOcc1);

                // 7. Mate ile YZ Plane (Soldan uzaklık)
                WorkPlane asmYZPlane = newAsmDef.WorkPlanes["YZ Plane"];
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc1, asmYZPlane, offsetXOcc1);

                // 8. Work Plane5'i al ve proxy'sini oluştur
                WorkPlane occWorkPlane5Occ1 = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc1;
                insertedOcc1.CreateGeometryProxy(occWorkPlane5Occ1, out workPlane5ProxyObjOcc1);
                WorkPlaneProxy workPlane5ProxyOcc1 = (WorkPlaneProxy)workPlane5ProxyObjOcc1;

                // 9. DirectedAngleConstraint ile 0 derece açı tanımla (yatay montaj)
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc1, asmXZPlane, 0);


                //OCC2
                ComponentOccurrence insertedOcc2 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ2XYPlane = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPlanes["XY Plane"];
                object occ2XYPlaneProxyObj;
                insertedOcc2.CreateGeometryProxy(occ2XYPlane, out occ2XYPlaneProxyObj);
                insertedOcc2.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ2XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc2 = 50.0 - GetParameterValue(insertedOcc2, "Length") / 2; // soldan uzaklık (mm)
                double offsetYOcc2 = GetParameterValue(insertedOcc2, "Length") + GetParameterValue(insertedOcc2, "Length") / 2; // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc2 = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc2;
                insertedOcc2.CreateGeometryProxy(workPtOcc2, out workPointProxyObjOcc2);
                WorkPointProxy workPointProxyOcc2 = (WorkPointProxy)workPointProxyObjOcc2;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc2, asmXZPlane, offsetYOcc2);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc2, asmYZPlane, offsetXOcc2);
                WorkPlane occWorkPlane5Occ2 = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc2;
                insertedOcc2.CreateGeometryProxy(occWorkPlane5Occ2, out workPlane5ProxyObjOcc2);
                WorkPlaneProxy workPlane5ProxyOcc2 = (WorkPlaneProxy)workPlane5ProxyObjOcc2;
                double angleOcc2InRadians = 90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc2, asmXZPlane, angleOcc2InRadians);


                //OCC3
                ComponentOccurrence insertedOcc3 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ3XYPlane = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPlanes["XY Plane"];
                object occ3XYPlaneProxyObj;
                insertedOcc3.CreateGeometryProxy(occ3XYPlane, out occ3XYPlaneProxyObj);
                insertedOcc3.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ3XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc3 = 50.0 + GetParameterValue(insertedOcc3, "Length") / 2; // soldan uzaklık (mm)
                double offsetYOcc3 = GetParameterValue(insertedOcc3, "Length") + GetParameterValue(insertedOcc3, "Length") / 2; // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc3 = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc3;
                insertedOcc3.CreateGeometryProxy(workPtOcc3, out workPointProxyObjOcc3);
                WorkPointProxy workPointProxyOcc3 = (WorkPointProxy)workPointProxyObjOcc3;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc3, asmXZPlane, offsetYOcc3);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc3, asmYZPlane, offsetXOcc3);
                WorkPlane occWorkPlane5Occ3 = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc3;
                insertedOcc3.CreateGeometryProxy(occWorkPlane5Occ3, out workPlane5ProxyObjOcc3);
                WorkPlaneProxy workPlane5ProxyOcc3 = (WorkPlaneProxy)workPlane5ProxyObjOcc3;
                double angleOcc3InRadians = -90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc3, asmXZPlane, angleOcc3InRadians);


                //OCC4
                ComponentOccurrence insertedOcc4 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ4XYPlane = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPlanes["XY Plane"];
                object occ4XYPlaneProxyObj;
                insertedOcc4.CreateGeometryProxy(occ4XYPlane, out occ4XYPlaneProxyObj);
                insertedOcc4.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ4XYPlaneProxyObj, asmXYPlane, 0);
                SetTextParameterValue(insertedOcc4, "Water_Discharge1", "None");
                SetTextParameterValue(insertedOcc4, "Water_Discharge2", "None");
                double offsetXOcc4 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc4 = GetParameterValue(insertedOcc4, "Length") + GetParameterValue(insertedOcc4, "Length"); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc4 = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc4;
                insertedOcc4.CreateGeometryProxy(workPtOcc4, out workPointProxyObjOcc4);
                WorkPointProxy workPointProxyOcc4 = (WorkPointProxy)workPointProxyObjOcc4;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc4, asmXZPlane, offsetYOcc4);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc4, asmYZPlane, offsetXOcc4);
                WorkPlane occWorkPlane5Occ4 = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc4;
                insertedOcc4.CreateGeometryProxy(occWorkPlane5Occ4, out workPlane5ProxyObjOcc4);
                WorkPlaneProxy workPlane5ProxyOcc4 = (WorkPlaneProxy)workPlane5ProxyObjOcc4;
                double angleOcc4InRadians = 180 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc4, asmXZPlane, angleOcc4InRadians);

                //OCC5

                string assemblyToInsertPathVol2 = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol2\60ST15_Kasa.iam";
                if (!System.IO.File.Exists(assemblyToInsertPathVol2))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPathVol2, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ComponentOccurrence insertedOcc5 = newAsmDef.Occurrences.Add(assemblyToInsertPathVol2, insertMatrix);
                WorkPlane occ5XYPlane = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPlanes["XY Plane"];
                object occ5XYPlaneProxyObj;
                insertedOcc5.CreateGeometryProxy(occ5XYPlane, out occ5XYPlaneProxyObj);
                insertedOcc5.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ5XYPlaneProxyObj, asmXYPlane, 0);
                SetTextParameterValue(insertedOcc5, "Corner_A_Assembly_Type", "A2");
                SetTextParameterValue(insertedOcc5, "Corner_B_Assembly_Type", "B2");
                SetTextParameterValue(insertedOcc5, "Water_Discharge1", "None");
                SetTextParameterValue(insertedOcc5, "Water_Discharge2", "None");
                SetNumericParameterValue(insertedOcc5, "Length", 1575.9);
                //SetNumericParameterValue(insertedOcc5, "B2_Size", -1000);
                double offsetXOcc5 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc5 = GetParameterValue(insertedOcc1, "Length") + (GetParameterValue(insertedOcc1, "Length") / 2); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc5 = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc5;
                insertedOcc5.CreateGeometryProxy(workPtOcc5, out workPointProxyObjOcc5);
                WorkPointProxy workPointProxyOcc5 = (WorkPointProxy)workPointProxyObjOcc5;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc5, asmXZPlane, offsetYOcc5);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc5, asmYZPlane, offsetXOcc5);
                WorkPlane occWorkPlane5Occ5 = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc5;
                insertedOcc5.CreateGeometryProxy(occWorkPlane5Occ5, out workPlane5ProxyObjOcc5);
                WorkPlaneProxy workPlane5ProxyOcc5 = (WorkPlaneProxy)workPlane5ProxyObjOcc5;
                double angleOcc5InRadians = 180 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc5, asmXZPlane, angleOcc5InRadians);

                //OCC6
                string assemblyToInsertPathVol3 = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol3\60ST15_Kasa.iam";
                if (!System.IO.File.Exists(assemblyToInsertPathVol3))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPathVol3, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ComponentOccurrence insertedOcc6 = newAsmDef.Occurrences.Add(assemblyToInsertPathVol3, insertMatrix);
                WorkPlane occ6XYPlane = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPlanes["XY Plane"];
                object occ6XYPlaneProxyObj;
                insertedOcc6.CreateGeometryProxy(occ6XYPlane, out occ6XYPlaneProxyObj);
                insertedOcc6.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ6XYPlaneProxyObj, asmXYPlane, 0);
                SetTextParameterValue(insertedOcc6, "Corner_A_Assembly_Type", "A2");
                SetTextParameterValue(insertedOcc6, "Corner_B_Assembly_Type", "B2");
                SetTextParameterValue(insertedOcc6, "Water_Discharge1", "None");
                SetTextParameterValue(insertedOcc6, "Water_Discharge2", "None");
                SetNumericParameterValue(insertedOcc6, "Length", 800);
                double offsetXOcc6 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc6 = GetParameterValue(insertedOcc1, "Length") + (GetParameterValue(insertedOcc1, "Length") * 0.75); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc6 = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc6;
                insertedOcc6.CreateGeometryProxy(workPtOcc6, out workPointProxyObjOcc6);
                WorkPointProxy workPointProxyOcc6 = (WorkPointProxy)workPointProxyObjOcc6;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc6, asmXZPlane, offsetYOcc6);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc6, asmYZPlane, offsetXOcc6);
                WorkPlane occWorkPlane5Occ6 = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc6;
                insertedOcc6.CreateGeometryProxy(occWorkPlane5Occ6, out workPlane5ProxyObjOcc6);
                WorkPlaneProxy workPlane5ProxyOcc6 = (WorkPlaneProxy)workPlane5ProxyObjOcc6;
                double angleOcc6InRadians = 90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc6, asmXZPlane, angleOcc6InRadians);

                //OCC6
                string assemblyToInsertPathVol4 = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol4\60ST15_Kasa.iam";
                if (!System.IO.File.Exists(assemblyToInsertPathVol4))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPathVol4, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ComponentOccurrence insertedOcc7 = newAsmDef.Occurrences.Add(assemblyToInsertPathVol4, insertMatrix);
                WorkPlane occ7XYPlane = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPlanes["XY Plane"];
                object occ7XYPlaneProxyObj;
                insertedOcc7.CreateGeometryProxy(occ7XYPlane, out occ7XYPlaneProxyObj);
                insertedOcc7.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ7XYPlaneProxyObj, asmXYPlane, 0);
                SetTextParameterValue(insertedOcc7, "Corner_A_Assembly_Type", "A2");
                SetTextParameterValue(insertedOcc7, "Corner_B_Assembly_Type", "B2");
                SetTextParameterValue(insertedOcc7, "Water_Discharge1", "None");
                SetTextParameterValue(insertedOcc7, "Water_Discharge2", "None");
                SetNumericParameterValue(insertedOcc7, "Length", 800);
                double offsetXOcc7 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc7 = GetParameterValue(insertedOcc1, "Length") + (GetParameterValue(insertedOcc1, "Length") * 0.25); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc7 = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc7;
                insertedOcc7.CreateGeometryProxy(workPtOcc7, out workPointProxyObjOcc7);
                WorkPointProxy workPointProxyOcc7 = (WorkPointProxy)workPointProxyObjOcc7;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc7, asmXZPlane, offsetYOcc7);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc7, asmYZPlane, offsetXOcc7);
                WorkPlane occWorkPlane5Occ7 = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc7;
                insertedOcc7.CreateGeometryProxy(occWorkPlane5Occ7, out workPlane5ProxyObjOcc7);
                WorkPlaneProxy workPlane5ProxyOcc7 = (WorkPlaneProxy)workPlane5ProxyObjOcc7;
                double angleOcc7InRadians = 90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc7, asmXZPlane, angleOcc7InRadians);

                MessageBox.Show("Montaj başarıyla eklendi, hizalandı ve konumlandırıldı!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Yeni montaj dosyasına eklenirken hata oluştu:\n" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private double GetParameterValue(ComponentOccurrence occ, string parameterName)
        {
            dynamic definition = occ.Definition;

            // Hem Part hem Assembly ComponentDefinition desteklenir
            Parameters parameters = definition.Parameters;
            Parameter param = parameters[parameterName];

            return param.Value;
        }

        private void SetTextParameterValue(ComponentOccurrence occ, string parameterName, string newValue)
        {
            dynamic definition = occ.Definition;
            Parameters parameters = definition.Parameters;
            Parameter param = parameters[parameterName];
            param.Value = newValue;
        }

        private void SetNumericParameterValue(ComponentOccurrence occ, string parameterName, double valueInMm)
        {
            dynamic definition = occ.Definition;

            Parameters parameters = definition.Parameters;
            Parameter param = parameters[parameterName];

            // Inventor varsayılan birimi cm'dir, mm'den cm'ye çevir
            double valueInCm = valueInMm / 10.0;

            param.Value = valueInCm;
        }
        private void btnCreateAllParts_Click(object sender, EventArgs e)
        {
            try
            {
                string assemblyToInsertPath = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol1\60ST15_Kasa.iam";

                if (!System.IO.File.Exists(assemblyToInsertPath))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPath, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 1. Boş bir montaj oluştur
                AssemblyDocument newAsmDoc = (AssemblyDocument)_inventorApp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, "", true);
                AssemblyComponentDefinition newAsmDef = newAsmDoc.ComponentDefinition;

                // 2. Dosyayı montaj ortamına ekle
                Matrix insertMatrix = _inventorApp.TransientGeometry.CreateMatrix();
                ComponentOccurrence insertedOcc1 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);

                // 3. Profilin XY Plane'ini montajın XY Plane'iyle hizala (Flush)
                WorkPlane asmXYPlane = newAsmDef.WorkPlanes["XY Plane"];
                WorkPlane occ1XYPlane = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPlanes["XY Plane"];
                object occ1XYPlaneProxyObj;
                insertedOcc1.CreateGeometryProxy(occ1XYPlane, out occ1XYPlaneProxyObj);
                insertedOcc1.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ1XYPlaneProxyObj, asmXYPlane, 0);

                // 4. Kullanıcıdan offset değerleri (örnek değerler burada sabit)
                double offsetXOcc1 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc1 = GetParameterValue(insertedOcc1, "Length"); // aşağıdan yükseklik (mm)

                // 5. WorkPoint1'i al ve proxy'sini oluştur
                WorkPoint workPtOcc1 = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc1;
                insertedOcc1.CreateGeometryProxy(workPtOcc1, out workPointProxyObjOcc1);
                WorkPointProxy workPointProxyOcc1 = (WorkPointProxy)workPointProxyObjOcc1;

                // 6. Mate ile XZ Plane (Yükseklik)
                WorkPlane asmXZPlane = newAsmDef.WorkPlanes["XZ Plane"];
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc1, asmXZPlane, offsetYOcc1);

                // 7. Mate ile YZ Plane (Soldan uzaklık)
                WorkPlane asmYZPlane = newAsmDef.WorkPlanes["YZ Plane"];
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc1, asmYZPlane, offsetXOcc1);

                // 8. Work Plane5'i al ve proxy'sini oluştur
                WorkPlane occWorkPlane5Occ1 = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc1;
                insertedOcc1.CreateGeometryProxy(occWorkPlane5Occ1, out workPlane5ProxyObjOcc1);
                WorkPlaneProxy workPlane5ProxyOcc1 = (WorkPlaneProxy)workPlane5ProxyObjOcc1;

                // 9. DirectedAngleConstraint ile 0 derece açı tanımla (yatay montaj)
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc1, asmXZPlane, 0);


                //OCC2
                ComponentOccurrence insertedOcc2 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ2XYPlane = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPlanes["XY Plane"];
                object occ2XYPlaneProxyObj;
                insertedOcc2.CreateGeometryProxy(occ2XYPlane, out occ2XYPlaneProxyObj);
                insertedOcc2.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ2XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc2 = 50.0 - GetParameterValue(insertedOcc2, "Length") / 2; // soldan uzaklık (mm)
                double offsetYOcc2 = GetParameterValue(insertedOcc2, "Length") + GetParameterValue(insertedOcc2, "Length") / 2; // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc2 = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc2;
                insertedOcc2.CreateGeometryProxy(workPtOcc2, out workPointProxyObjOcc2);
                WorkPointProxy workPointProxyOcc2 = (WorkPointProxy)workPointProxyObjOcc2;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc2, asmXZPlane, offsetYOcc2);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc2, asmYZPlane, offsetXOcc2);
                WorkPlane occWorkPlane5Occ2 = ((AssemblyComponentDefinition)insertedOcc2.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc2;
                insertedOcc2.CreateGeometryProxy(occWorkPlane5Occ2, out workPlane5ProxyObjOcc2);
                WorkPlaneProxy workPlane5ProxyOcc2 = (WorkPlaneProxy)workPlane5ProxyObjOcc2;
                double angleOcc2InRadians = 90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc2, asmXZPlane, angleOcc2InRadians);


                //OCC3
                ComponentOccurrence insertedOcc3 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ3XYPlane = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPlanes["XY Plane"];
                object occ3XYPlaneProxyObj;
                insertedOcc3.CreateGeometryProxy(occ3XYPlane, out occ3XYPlaneProxyObj);
                insertedOcc3.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ3XYPlaneProxyObj, asmXYPlane, 0);
                double offsetXOcc3 = 50.0 + GetParameterValue(insertedOcc3, "Length") / 2; // soldan uzaklık (mm)
                double offsetYOcc3 = GetParameterValue(insertedOcc3, "Length") + GetParameterValue(insertedOcc3, "Length") / 2; // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc3 = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc3;
                insertedOcc3.CreateGeometryProxy(workPtOcc3, out workPointProxyObjOcc3);
                WorkPointProxy workPointProxyOcc3 = (WorkPointProxy)workPointProxyObjOcc3;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc3, asmXZPlane, offsetYOcc3);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc3, asmYZPlane, offsetXOcc3);
                WorkPlane occWorkPlane5Occ3 = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc3;
                insertedOcc3.CreateGeometryProxy(occWorkPlane5Occ3, out workPlane5ProxyObjOcc3);
                WorkPlaneProxy workPlane5ProxyOcc3 = (WorkPlaneProxy)workPlane5ProxyObjOcc3;
                double angleOcc3InRadians = -90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc3, asmXZPlane, angleOcc3InRadians);


                //OCC4
                ComponentOccurrence insertedOcc4 = newAsmDef.Occurrences.Add(assemblyToInsertPath, insertMatrix);
                WorkPlane occ4XYPlane = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPlanes["XY Plane"];
                object occ4XYPlaneProxyObj;
                insertedOcc4.CreateGeometryProxy(occ4XYPlane, out occ4XYPlaneProxyObj);
                insertedOcc4.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ4XYPlaneProxyObj, asmXYPlane, 0);
                SetTextParameterValue(insertedOcc4, "Water_Discharge1", "None");
                SetTextParameterValue(insertedOcc4, "Water_Discharge2", "None");
                double offsetXOcc4 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc4 = GetParameterValue(insertedOcc4, "Length") + GetParameterValue(insertedOcc4, "Length"); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc4 = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc4;
                insertedOcc4.CreateGeometryProxy(workPtOcc4, out workPointProxyObjOcc4);
                WorkPointProxy workPointProxyOcc4 = (WorkPointProxy)workPointProxyObjOcc4;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc4, asmXZPlane, offsetYOcc4);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc4, asmYZPlane, offsetXOcc4);
                WorkPlane occWorkPlane5Occ4 = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc4;
                insertedOcc4.CreateGeometryProxy(occWorkPlane5Occ4, out workPlane5ProxyObjOcc4);
                WorkPlaneProxy workPlane5ProxyOcc4 = (WorkPlaneProxy)workPlane5ProxyObjOcc4;
                double angleOcc4InRadians = 180 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc4, asmXZPlane, angleOcc4InRadians);


                //OCC5

                string assemblyToInsertPathVol2 = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol2\60ST15_Kasa.iam";
                if (!System.IO.File.Exists(assemblyToInsertPathVol2))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPathVol2, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ComponentOccurrence insertedOcc5 = newAsmDef.Occurrences.Add(assemblyToInsertPathVol2, insertMatrix);
                WorkPlane occ5XYPlane = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPlanes["XY Plane"];
                object occ5XYPlaneProxyObj;
                insertedOcc5.CreateGeometryProxy(occ5XYPlane, out occ5XYPlaneProxyObj);
                insertedOcc5.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ5XYPlaneProxyObj, asmXYPlane, 0);
                SetTextParameterValue(insertedOcc5, "Corner_A_Assembly_Type", "A2");
                SetTextParameterValue(insertedOcc5, "Corner_B_Assembly_Type", "B2");
                SetTextParameterValue(insertedOcc5, "Water_Discharge1", "None");
                SetTextParameterValue(insertedOcc5, "Water_Discharge2", "None");
                SetNumericParameterValue(insertedOcc5, "Length", 1575.9);
                //SetNumericParameterValue(insertedOcc5, "B2_Size", -1000);
                double offsetXOcc5 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc5 = GetParameterValue(insertedOcc1, "Length") + (GetParameterValue(insertedOcc1, "Length") / 2); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc5 = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc5;
                insertedOcc5.CreateGeometryProxy(workPtOcc5, out workPointProxyObjOcc5);
                WorkPointProxy workPointProxyOcc5 = (WorkPointProxy)workPointProxyObjOcc5;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc5, asmXZPlane, offsetYOcc5);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc5, asmYZPlane, offsetXOcc5);
                WorkPlane occWorkPlane5Occ5 = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc5;
                insertedOcc5.CreateGeometryProxy(occWorkPlane5Occ5, out workPlane5ProxyObjOcc5);
                WorkPlaneProxy workPlane5ProxyOcc5 = (WorkPlaneProxy)workPlane5ProxyObjOcc5;
                double angleOcc5InRadians = 180 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc5, asmXZPlane, angleOcc5InRadians);

                //OCC6
                string assemblyToInsertPathVol3 = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol3\60ST15_Kasa.iam";
                if (!System.IO.File.Exists(assemblyToInsertPathVol3))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPathVol3, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ComponentOccurrence insertedOcc6 = newAsmDef.Occurrences.Add(assemblyToInsertPathVol3, insertMatrix);
                WorkPlane occ6XYPlane = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPlanes["XY Plane"];
                object occ6XYPlaneProxyObj;
                insertedOcc6.CreateGeometryProxy(occ6XYPlane, out occ6XYPlaneProxyObj);
                insertedOcc6.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ6XYPlaneProxyObj, asmXYPlane, 0);
                SetTextParameterValue(insertedOcc6, "Corner_A_Assembly_Type", "A2");
                SetTextParameterValue(insertedOcc6, "Corner_B_Assembly_Type", "B2");
                SetTextParameterValue(insertedOcc6, "Water_Discharge1", "None");
                SetTextParameterValue(insertedOcc6, "Water_Discharge2", "None");
                SetNumericParameterValue(insertedOcc6, "Length", 800);
                double offsetXOcc6 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc6 = GetParameterValue(insertedOcc1, "Length") + (GetParameterValue(insertedOcc1, "Length") * 0.75); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc6 = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc6;
                insertedOcc6.CreateGeometryProxy(workPtOcc6, out workPointProxyObjOcc6);
                WorkPointProxy workPointProxyOcc6 = (WorkPointProxy)workPointProxyObjOcc6;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc6, asmXZPlane, offsetYOcc6);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc6, asmYZPlane, offsetXOcc6);
                WorkPlane occWorkPlane5Occ6 = ((AssemblyComponentDefinition)insertedOcc6.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc6;
                insertedOcc6.CreateGeometryProxy(occWorkPlane5Occ6, out workPlane5ProxyObjOcc6);
                WorkPlaneProxy workPlane5ProxyOcc6 = (WorkPlaneProxy)workPlane5ProxyObjOcc6;
                double angleOcc6InRadians = 90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc6, asmXZPlane, angleOcc6InRadians);

                //OCC6
                string assemblyToInsertPathVol4 = @"C:\Users\mustafa yucel\Desktop\Montaj\60ST15Vol4\60ST15_Kasa.iam";
                if (!System.IO.File.Exists(assemblyToInsertPathVol4))
                {
                    MessageBox.Show("Eklenmek istenen montaj dosyası bulunamadı:\n" + assemblyToInsertPathVol4, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ComponentOccurrence insertedOcc7 = newAsmDef.Occurrences.Add(assemblyToInsertPathVol4, insertMatrix);
                WorkPlane occ7XYPlane = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPlanes["XY Plane"];
                object occ7XYPlaneProxyObj;
                insertedOcc7.CreateGeometryProxy(occ7XYPlane, out occ7XYPlaneProxyObj);
                insertedOcc7.Grounded = false;
                newAsmDef.Constraints.AddFlushConstraint((WorkPlaneProxy)occ7XYPlaneProxyObj, asmXYPlane, 0);
                SetTextParameterValue(insertedOcc7, "Corner_A_Assembly_Type", "A2");
                SetTextParameterValue(insertedOcc7, "Corner_B_Assembly_Type", "B2");
                SetTextParameterValue(insertedOcc7, "Water_Discharge1", "None");
                SetTextParameterValue(insertedOcc7, "Water_Discharge2", "None");
                SetNumericParameterValue(insertedOcc7, "Length", 800);
                double offsetXOcc7 = 50.0; // soldan uzaklık (mm)
                double offsetYOcc7 = GetParameterValue(insertedOcc1, "Length") + (GetParameterValue(insertedOcc1, "Length") * 0.25); // aşağıdan yükseklik (mm)
                WorkPoint workPtOcc7 = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPoints["Work Point1"];
                object workPointProxyObjOcc7;
                insertedOcc7.CreateGeometryProxy(workPtOcc7, out workPointProxyObjOcc7);
                WorkPointProxy workPointProxyOcc7 = (WorkPointProxy)workPointProxyObjOcc7;
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc7, asmXZPlane, offsetYOcc7);
                newAsmDef.Constraints.AddMateConstraint(workPointProxyOcc7, asmYZPlane, offsetXOcc7);
                WorkPlane occWorkPlane5Occ7 = ((AssemblyComponentDefinition)insertedOcc7.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjOcc7;
                insertedOcc7.CreateGeometryProxy(occWorkPlane5Occ7, out workPlane5ProxyObjOcc7);
                WorkPlaneProxy workPlane5ProxyOcc7 = (WorkPlaneProxy)workPlane5ProxyObjOcc7;
                double angleOcc7InRadians = 90 * Math.PI / 180;
                newAsmDef.Constraints.AddAngleConstraint(workPlane5ProxyOcc7, asmXZPlane, angleOcc7InRadians);

                //CAM 1 MONTAGE
                string camMontageToInsertPathVol1 = @"C:\Users\mustafa yucel\Desktop\Montaj\GlassMontage1\4ddb326a-df43-4679-9ed4-cf2741b6951b.iam";
                if (!System.IO.File.Exists(camMontageToInsertPathVol1))
                {
                    MessageBox.Show("Eklenmek istenen cam dosyası bulunamadı:\n" + camMontageToInsertPathVol1, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ComponentOccurrence insertedCam1 = newAsmDef.Occurrences.Add(camMontageToInsertPathVol1, insertMatrix);
                ComponentOccurrence innerPartOcc1 = insertedCam1.SubOccurrences.Cast<ComponentOccurrence>().FirstOrDefault(occ => occ.Name.Contains("171988f7-8dd0-4b03-adb8-bcefea351882"));
                if (innerPartOcc1 == null)
                {
                    MessageBox.Show("İstenen parça bulunamadı.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                PartComponentDefinition partDef1 = (PartComponentDefinition)innerPartOcc1.Definition;
                WorkPlane cam1HWP0Plane = partDef1.WorkPlanes["HWP0"];
                object cam1HWP0PlaneProxyObj;
                innerPartOcc1.CreateGeometryProxy(cam1HWP0Plane, out cam1HWP0PlaneProxyObj);
                WorkPlaneProxy cam1HWP0PlaneProxy = (WorkPlaneProxy)cam1HWP0PlaneProxyObj;
                insertedCam1.Grounded = false;
                WorkPlane occWorkPlane5Cam1 = ((AssemblyComponentDefinition)insertedOcc1.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjCam1;
                insertedOcc1.CreateGeometryProxy(occWorkPlane5Cam1, out workPlane5ProxyObjCam1);
                WorkPlaneProxy workPlane5ProxyCam1 = (WorkPlaneProxy)workPlane5ProxyObjCam1;
                newAsmDef.Constraints.AddMateConstraint(workPlane5ProxyCam1, cam1HWP0PlaneProxy, 0);
                insertedCam1.Grounded = true;

                //CAM 2 MONTAGE
                string camMontageToInsertPathVol2 = @"C:\Users\mustafa yucel\Desktop\Montaj\GlassMontage2\4ddb326a-df43-4679-9ed4-cf2741b6951b.iam";
                if (!System.IO.File.Exists(camMontageToInsertPathVol2))
                {
                    MessageBox.Show("Eklenmek istenen cam dosyası bulunamadı:\n" + camMontageToInsertPathVol2, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ComponentOccurrence insertedCam2 = newAsmDef.Occurrences.Add(camMontageToInsertPathVol2, insertMatrix);
                ComponentOccurrence innerPartOcc2 = insertedCam2.SubOccurrences.Cast<ComponentOccurrence>().FirstOrDefault(occ => occ.Name.Contains("171988f7-8dd0-4b03-adb8-bcefea351882"));
                if (innerPartOcc2 == null)
                {
                    MessageBox.Show("İstenen parça bulunamadı.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                PartComponentDefinition partDef2 = (PartComponentDefinition)innerPartOcc2.Definition;
                WorkPlane cam2HWP0Plane = partDef2.WorkPlanes["HWP2"];
                object cam2HWP0PlaneProxyObj;
                innerPartOcc2.CreateGeometryProxy(cam2HWP0Plane, out cam2HWP0PlaneProxyObj);
                WorkPlaneProxy cam2HWP0PlaneProxy = (WorkPlaneProxy)cam2HWP0PlaneProxyObj;
                insertedCam2.Grounded = false;
                WorkPlane occWorkPlane5Cam2 = ((AssemblyComponentDefinition)insertedOcc4.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjCam2;
                insertedOcc4.CreateGeometryProxy(occWorkPlane5Cam2, out workPlane5ProxyObjCam2);
                WorkPlaneProxy workPlane5ProxyCam2 = (WorkPlaneProxy)workPlane5ProxyObjCam2;
                newAsmDef.Constraints.AddMateConstraint(workPlane5ProxyCam2, cam2HWP0PlaneProxy, 0);
                insertedCam2.Grounded = true;

                //CAM 3 MONTAGE
                string camMontageToInsertPathVol3 = @"C:\Users\mustafa yucel\Desktop\Montaj\GlassMontage3\4ddb326a-df43-4679-9ed4-cf2741b6951b.iam";
                if (!System.IO.File.Exists(camMontageToInsertPathVol3))
                {
                    MessageBox.Show("Eklenmek istenen cam dosyası bulunamadı:\n" + camMontageToInsertPathVol3, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ComponentOccurrence insertedCam3 = newAsmDef.Occurrences.Add(camMontageToInsertPathVol3, insertMatrix);
                ComponentOccurrence innerPartOcc3 = insertedCam3.SubOccurrences.Cast<ComponentOccurrence>().FirstOrDefault(occ => occ.Name.Contains("171988f7-8dd0-4b03-adb8-bcefea351882"));
                if (innerPartOcc3 == null)
                {
                    MessageBox.Show("İstenen parça bulunamadı.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                PartComponentDefinition partDef3 = (PartComponentDefinition)innerPartOcc3.Definition;
                WorkPlane cam3HWP0Plane = partDef3.WorkPlanes["HWP1"];
                object cam3HWP0PlaneProxyObj;
                innerPartOcc3.CreateGeometryProxy(cam3HWP0Plane, out cam3HWP0PlaneProxyObj);
                WorkPlaneProxy cam3HWP0PlaneProxy = (WorkPlaneProxy)cam3HWP0PlaneProxyObj;
                insertedCam3.Grounded = false;
                WorkPlane occWorkPlane5Cam3 = ((AssemblyComponentDefinition)insertedOcc3.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjCam3;
                insertedOcc3.CreateGeometryProxy(occWorkPlane5Cam3, out workPlane5ProxyObjCam3);
                WorkPlaneProxy workPlane5ProxyCam3 = (WorkPlaneProxy)workPlane5ProxyObjCam3;
                newAsmDef.Constraints.AddMateConstraint(workPlane5ProxyCam3, cam3HWP0PlaneProxy, 0);
                insertedCam3.Grounded = true;


                //CAM 4 MONTAGE
                string camMontageToInsertPathVol4 = @"C:\Users\mustafa yucel\Desktop\Montaj\GlassMontage4\4ddb326a-df43-4679-9ed4-cf2741b6951b.iam";
                if (!System.IO.File.Exists(camMontageToInsertPathVol4))
                {
                    MessageBox.Show("Eklenmek istenen cam dosyası bulunamadı:\n" + camMontageToInsertPathVol4, "Dosya Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ComponentOccurrence insertedCam4 = newAsmDef.Occurrences.Add(camMontageToInsertPathVol4, insertMatrix);
                ComponentOccurrence innerPartOcc4 = insertedCam4.SubOccurrences.Cast<ComponentOccurrence>().FirstOrDefault(occ => occ.Name.Contains("171988f7-8dd0-4b03-adb8-bcefea351882"));
                if (innerPartOcc4 == null)
                {
                    MessageBox.Show("İstenen parça bulunamadı.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                PartComponentDefinition partDef4 = (PartComponentDefinition)innerPartOcc4.Definition;
                WorkPlane cam4HWP0Plane = partDef4.WorkPlanes["HWP2"];
                object cam4HWP0PlaneProxyObj;
                innerPartOcc4.CreateGeometryProxy(cam4HWP0Plane, out cam4HWP0PlaneProxyObj);
                WorkPlaneProxy cam4HWP0PlaneProxy = (WorkPlaneProxy)cam4HWP0PlaneProxyObj;
                insertedCam4.Grounded = false;
                WorkPlane occWorkPlane5Cam4 = ((AssemblyComponentDefinition)insertedOcc5.Definition).WorkPlanes["Work Plane5"];
                object workPlane5ProxyObjCam4;
                insertedOcc5.CreateGeometryProxy(occWorkPlane5Cam4, out workPlane5ProxyObjCam4);
                WorkPlaneProxy workPlane5ProxyCam4 = (WorkPlaneProxy)workPlane5ProxyObjCam4;
                newAsmDef.Constraints.AddMateConstraint(workPlane5ProxyCam4, cam4HWP0PlaneProxy, 0);
                insertedCam4.Grounded = true;

                MessageBox.Show("Montaj başarıyla eklendi, hizalandı ve konumlandırıldı!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Yeni montaj dosyasına eklenirken hata oluştu:\n" + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
    }
}
