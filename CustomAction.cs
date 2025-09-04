using System;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using WixToolset.Dtf.WindowsInstaller;

namespace VfcPatcher.CA
{
    public class CustomActions
    {
        // -------- Utility: Version detection --------
        static bool TryGetFileVersion(string path, out Version v)
        {
            v = null;
            try
            {
                var f = FileVersionInfo.GetVersionInfo(path);
                if (f.FileMajorPart < 0 || f.FileMinorPart < 0) return false;

                if (f.FileBuildPart >= 0 && f.FilePrivatePart >= 0)
                    v = new Version(f.FileMajorPart, f.FileMinorPart, f.FileBuildPart, f.FilePrivatePart);
                else if (f.FileBuildPart >= 0)
                    v = new Version(f.FileMajorPart, f.FileMinorPart, f.FileBuildPart);
                else
                    v = new Version(f.FileMajorPart, f.FileMinorPart);

                return true;
            }
            catch { return false; }
        }

        // -------- Utility: Wait until a file is unlocked --------
        static void WaitForFileRelease(string file, Session session, int delayMs = 1000)
        {
            while (true)
            {
                try
                {
                    using (FileStream fs = File.Open(file, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        session.Log("File is no longer locked: {0}", file);
                        return;
                    }
                }
                catch (IOException)
                {
                    session.Log("File still locked, waiting... {0}", file);
                    Thread.Sleep(delayMs);
                }
            }
        }

        // -------- Service control helpers --------
        static void StopService(string name, Session session, int timeoutSec = 60)
        {
            using (var sc = new ServiceController(name))
            {
                session.Log("Checking status of {0}", name);
                if (sc.Status != ServiceControllerStatus.Stopped && sc.Status != ServiceControllerStatus.StopPending)
                {
                    sc.Stop();
                    sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(timeoutSec));
                    session.Log("{0} stopped", name);
                }
                else session.Log("{0} already stopped or stopping", name);
            }
        }

        static void StartService(string name, Session session, int timeoutSec = 60)
        {
            using (var sc = new ServiceController(name))
            {
                session.Log("Checking status of {0}", name);
                if (sc.Status != ServiceControllerStatus.Running && sc.Status != ServiceControllerStatus.StartPending)
                {
                    sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(timeoutSec));
                    session.Log("{0} started", name);
                }
                else session.Log("{0} already running or starting", name);
            }
        }

        // -------- CSV helpers --------
        static void AppendCsv(string path, string[] values)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            bool newFile = !File.Exists(path);
            using (var sw = new StreamWriter(path, append: true, new UTF8Encoding(false)))
            {
                if (newFile)
                    sw.WriteLine("Date,Patch No,Service Version,Base Version,Patched Version,Patch Status");

                sw.WriteLine(string.Join(",", Array.ConvertAll(values, CsvEscape)));
            }
        }

        static string CsvEscape(string s) =>
            (s.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0) ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;

        // ================= Custom Actions =================

        [CustomAction]
        public static ActionResult CA_Precheck(Session session)
        {
            try
            {
                var data = session.CustomActionData;

                string targetExe = data["TARGETEXE"];
                string patchVerS = data["PATCHVERSION"];
                string regBaseV = data["REGBASEVER"];

                if (!Version.TryParse(patchVerS, out var patchVer))
                    throw new InstallException("Invalid patch version format.");

                if (!File.Exists(targetExe))
                    throw new InstallException(string.Format("File not found: {0}", targetExe));

                if (!TryGetFileVersion(targetExe, out var existingVer))
                    throw new InstallException("Unable to read current service file version.");

                if (!Version.TryParse(regBaseV, out var baseVer))
                    throw new InstallException("Unable to read base version from registry.");

                if (patchVer < baseVer)
                    throw new InstallException("Patching Failed. Base Version of VFC is higher than the Current Patch Version");

                if (patchVer <= existingVer)
                    throw new InstallException("Higher or Same Version of Verba Storage Service is already running, Service cannot be Patched");

                session.Log("Prechecks passed. Base={0}, Existing={1}, Patch={2}", baseVer, existingVer, patchVer);
                return ActionResult.Success;
            }
            catch (Exception ex)
            {
                session.Log("ERROR in CA_Precheck: {0}", ex);
                return ActionResult.Failure;
            }
        }

        [CustomAction]
        public static ActionResult CA_Patch(Session session)
        {
            try
            {
                var data = session.CustomActionData;

                string patchVerS = data["PATCHVERSION"];
                string targetExe = data["TARGETEXE"];
                string payload = data["PAYLOAD"];
                string logDir = data["LOGDIR"];
                string svcPatch = data["SVCPATCH"];
                string svcSysMon = data["SVCSYSMON"];
                string regBaseV = data["REGBASEVER"];

                Directory.CreateDirectory(logDir);

                Version.TryParse(patchVerS, out var patchVer);
                TryGetFileVersion(targetExe, out var preServiceVer);
                Version.TryParse(regBaseV, out var baseVer);

                session.Log("Patch starting: Base={0}, Service={1}, Patch={2}", baseVer, preServiceVer, patchVer);

                // Disable services so they don't auto-start during replacement
                RunHidden(session, $"sc config \"{svcSysMon}\" start= disabled");
                RunHidden(session, $"sc config \"{svcPatch}\" start= disabled");

                StopService(svcSysMon, session);
                StopService(svcPatch, session);

                WaitForFileRelease(targetExe, session);

                // Backup current exe
                string backupRoot = Path.Combine(logDir, "Patch Backups");
                Directory.CreateDirectory(backupRoot);
                string backupName = $"verbastorage_{preServiceVer}.exe";
                string backupPath = EnsureUniquePath(Path.Combine(backupRoot, backupName));

                // Clear read-only bit if present (System.IO.FileAttributes!)
                try
                {
                    var attr = File.GetAttributes(targetExe);
                    if ((attr & System.IO.FileAttributes.ReadOnly) != 0)
                    {
                        File.SetAttributes(targetExe, attr & ~System.IO.FileAttributes.ReadOnly);
                    }
                }
                catch { /* best-effort */ }

                File.Move(targetExe, backupPath);
                session.Log("Backed up to: {0}", backupPath);

                File.Copy(payload, targetExe, overwrite: true);
                session.Log("Replaced target with payload: {0}", payload);

                // Restore service startup and start
                RunHidden(session, $"sc config \"{svcPatch}\" start= auto");
                RunHidden(session, $"sc config \"{svcSysMon}\" start= auto");

                StartService(svcPatch, session);
                StartService(svcSysMon, session);

                try { AppendXlsxAudit(logDir, "Success", patchVer, baseVer, preServiceVer); }
                catch { AppendCsvAudit(logDir, "Success", patchVer, baseVer, preServiceVer); }

                return ActionResult.Success;
            }
            catch (Exception ex)
            {
                try
                {
                    var data = session.CustomActionData;
                    Version.TryParse(data["PATCHVERSION"], out var patchVer);
                    Version.TryParse(data["REGBASEVER"], out var baseVer);
                    TryGetFileVersion(data["TARGETEXE"], out var preServiceVer);
                    string logDir = data["LOGDIR"];
                    AppendCsvAudit(logDir, "Failed", patchVer, baseVer, preServiceVer);
                }
                catch { /* swallow secondary logging errors */ }

                session.Log("ERROR in CA_Patch: {0}", ex);
                return ActionResult.Failure;
            }
        }

        [CustomAction]
        public static ActionResult CA_Rollback(Session session)
        {
            try
            {
                session.Log("Rollback handler executed.");
                return ActionResult.Success;
            }
            catch (Exception ex)
            {
                session.Log("ERROR in CA_Rollback: {0}", ex);
                return ActionResult.Success; // be tolerant on rollback
            }
        }

        // -------- Shell helper --------
        static void RunHidden(Session s, string cmd)
        {
            var psi = new ProcessStartInfo("cmd.exe", "/c " + cmd)
            {
                CreateNoWindow = true,
                UseShellExecute = false
            };

            using (var p = Process.Start(psi))
            {
                if (p != null) p.WaitForExit();
            }

            s.Log("Executed: {0}", cmd);
        }

        // -------- File naming helper --------
        static string EnsureUniquePath(string path)
        {
            if (!File.Exists(path)) return path;

            var dir = Path.GetDirectoryName(path)!;
            var name = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);

            string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            int i = 2;
            string candidate = Path.Combine(dir, $"{name}_{ts}{ext}");
            while (File.Exists(candidate))
                candidate = Path.Combine(dir, $"{name}_{ts}_{i++}{ext}");

            return candidate;
        }

        // -------- XLSX audit (preferred) --------
        static void AppendXlsxAudit(string logDir, string status, Version patched, Version baseV, Version preServiceV)
        {
            string xlsx = Path.Combine(logDir, "VFC Patches.xlsx");
            if (!File.Exists(xlsx)) CreateWorkbookWithHeader(xlsx);

            using (var doc = SpreadsheetDocument.Open(xlsx, true))
            {
                var wb = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart missing");
                var ws = GetOrCreateFirstWorksheet(wb, "Patches");

                int nextNo = GetDataRowCount(ws) + 1;

                var sd = ws.Worksheet.GetFirstChild<SheetData>();
                var row = new Row();

                row.Append(
                    MakeTextCell(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
                    MakeTextCell(nextNo.ToString()),
                    MakeTextCell(preServiceV != null ? preServiceV.ToString() : string.Empty),
                    MakeTextCell(baseV != null ? baseV.ToString() : string.Empty),
                    MakeTextCell(patched != null ? patched.ToString() : string.Empty),
                    MakeTextCell(status)
                );

                sd.Append(row);
                ws.Worksheet.Save();
                wb.Workbook.Save();
            }
        }

        // -------- CSV audit (fallback) --------
        static void AppendCsvAudit(string logDir, string status, Version patched, Version baseV, Version preServiceV)
        {
            string csv = Path.Combine(logDir, "VFC Patches.csv");
            int count = 0;

            if (File.Exists(csv))
            {
                using (var sr = new StreamReader(csv))
                {
                    sr.ReadLine(); // header
                    while (true)
                    {
                        string line = sr.ReadLine();
                        if (line == null) break;
                        if (!string.IsNullOrWhiteSpace(line)) count++;
                    }
                }
            }

            AppendCsv(csv, new[]
            {
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                (count + 1).ToString(),
                preServiceV != null ? preServiceV.ToString() : "",
                baseV != null ? baseV.ToString() : "",
                patched != null ? patched.ToString() : "",
                status
            });
        }

        // -------- OpenXML helpers --------
        static void CreateWorkbookWithHeader(string path)
        {
            using (var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var wb = doc.AddWorkbookPart();
                wb.Workbook = new Workbook();

                var ws = wb.AddNewPart<WorksheetPart>();
                ws.Worksheet = new Worksheet(new SheetData());

                var sheets = wb.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Patches" };
                sheets.Append(sheet);

                var sd = ws.Worksheet.GetFirstChild<SheetData>();
                var header = new Row();
                string[] cols = { "Date", "Patch No", "Service Version", "Base Version", "Patched Version", "Patch Status" };
                foreach (var c in cols) header.Append(MakeTextCell(c));
                sd.Append(header);

                ws.Worksheet.Save();
                wb.Workbook.Save();
            }
        }

        static WorksheetPart GetOrCreateFirstWorksheet(WorkbookPart wb, string name)
        {
            var sheets = wb.Workbook.Sheets ?? wb.Workbook.AppendChild(new Sheets());
            if (!sheets.Elements<Sheet>().Any())
            {
                var ws = wb.AddNewPart<WorksheetPart>();
                ws.Worksheet = new Worksheet(new SheetData());
                sheets.Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = name });
                ws.Worksheet.Save(); wb.Workbook.Save();
                return ws;
            }
            else
            {
                var first = sheets.Elements<Sheet>().First();
                return (WorksheetPart)wb.GetPartById(first.Id);
            }
        }

        static int GetDataRowCount(WorksheetPart ws)
        {
            var sd = ws.Worksheet.GetFirstChild<SheetData>();
            int rows = 0;
            foreach (var _ in sd.Elements<Row>()) rows++;
            return Math.Max(0, rows - 1); // exclude header
        }

        static Cell MakeTextCell(string text)
        {
            return new Cell
            {
                DataType = CellValues.String,
                CellValue = new CellValue(text)
            };
        }
    }
}
