using System;
using System.IO;
using System.Security.Cryptography;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace HMVTools
{
    // ═══════════════════════════════════════════════════════════════
    //  Plain data class — no Revit references, safe for WPF windows
    // ═══════════════════════════════════════════════════════════════
    public class TraceData
    {
        public string AdcPath { get; set; } = "";
        public string Sha256Hash { get; set; } = "";
        public string LoadTimestamp { get; set; } = "";
        public string UserMachineId { get; set; } = "";
    }

    // ═══════════════════════════════════════════════════════════════
    //  Core traceability logic: ExtensibleStorage, hashing, hydration
    // ═══════════════════════════════════════════════════════════════

    /// <summary>
    /// Handles all file-system and Revit-side operations for family
    /// traceability: ADC file hydration, SHA-256 hashing, Extensible
    /// Storage read/write, and visible parameter injection.
    /// </summary>
    public static class FamilyTraceabilityManager
    {
        // ── Schema constants ──────────────────────────────────────
        // IMPORTANT: Never change this GUID after the first deployment.
        // Doing so would orphan data already written to .rvt files.
        private static readonly Guid SchemaGuid =
            new Guid("8F3A72C1-4E5B-4D9A-B6F0-1C2D3E4F5A6B");
        private const string SCHEMA_NAME = "HMV_FamilyTraceability";
        private const string F_ADC_PATH = "AdcPath";
        private const string F_SHA256 = "Sha256Hash";
        private const string F_TIMESTAMP = "LoadTimestamp";
        private const string F_USER_MACH = "UserMachineId";

        /// <summary>Name of the visible Type Parameter injected into families.</summary>
        public const string PARAM_ADC_PATH = "ADC_Library_Path";

        // ── ADC detection ─────────────────────────────────────────
        // Desktop Connector syncs to: C:\Users\<user>\DC\ACCDocs\...
        private static readonly string ADC_MARKER =
            Path.Combine("DC", "ACCDocs").ToLowerInvariant();

        /// <summary>Returns true if <paramref name="path"/> lives inside the ADC workspace.</summary>
        public static bool IsAdcPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            return path.ToLowerInvariant().Contains(ADC_MARKER);
        }

        // ══════════════════════════════════════════════════════════
        //  ADC FILE HYDRATION
        // ══════════════════════════════════════════════════════════

        /// <summary>
        /// Ensures a ProjFS-virtualized file is fully downloaded (hydrated)
        /// before any Revit API or crypto operations touch it.
        /// 
        /// ADC v16+ uses Windows Projected File System (ProjFS). Files
        /// marked "Online Only" have <see cref="FileAttributes.Offline"/>
        /// and are 0 bytes until an I/O read triggers hydration.
        /// 
        /// We open a read-only, non-locking <see cref="FileStream"/> and
        /// read a single byte — this is the minimum trigger Windows needs
        /// to download the file from ACC.
        /// </summary>
        /// <returns>True if the file is ready to read; false on error.</returns>
        public static bool EnsureHydrated(string path)
        {
            try
            {
                if (!File.Exists(path)) return false;

                FileAttributes attr = File.GetAttributes(path);
                bool isOffline = (attr & FileAttributes.Offline) == FileAttributes.Offline;

                if (isOffline)
                {
                    // CRITICAL: FileShare.ReadWrite — never lock ADC-managed files.
                    // Reading one byte is enough to trigger ProjFS hydration.
                    using (FileStream fs = new FileStream(
                        path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        fs.ReadByte();
                    }
                }
                return true;
            }
            catch (IOException) { return false; }
            catch (UnauthorizedAccessException) { return false; }
        }

        // ══════════════════════════════════════════════════════════
        //  SHA-256 HASHING
        // ══════════════════════════════════════════════════════════

        /// <summary>
        /// Computes the SHA-256 hash of a file using read-only, non-locking access.
        /// Call <see cref="EnsureHydrated"/> first for ADC paths.
        /// </summary>
        /// <returns>Lowercase hex string, or null on failure.</returns>
        public static string ComputeSha256(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(
                    path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (SHA256 sha = SHA256.Create())
                {
                    byte[] hash = sha.ComputeHash(fs);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
            catch (Exception) { return null; }
        }

        // ══════════════════════════════════════════════════════════
        //  EXTENSIBLE STORAGE — Schema
        // ══════════════════════════════════════════════════════════

        /// <summary>
        /// Retrieves the existing schema or creates a new one.
        /// Schema fields: AdcPath, Sha256Hash, LoadTimestamp, UserMachineId.
        /// </summary>
        private static Schema GetOrCreateSchema()
        {
            Schema schema = Schema.Lookup(SchemaGuid);
            if (schema != null) return schema;

            SchemaBuilder sb = new SchemaBuilder(SchemaGuid);
            sb.SetSchemaName(SCHEMA_NAME);
            sb.SetReadAccessLevel(AccessLevel.Public);
            sb.SetWriteAccessLevel(AccessLevel.Public);
            // Vendor ID must match .addin manifest. Using "HMV" as identifier.
            sb.SetVendorId("HMV");

            sb.AddSimpleField(F_ADC_PATH, typeof(string));
            sb.AddSimpleField(F_SHA256, typeof(string));
            sb.AddSimpleField(F_TIMESTAMP, typeof(string));
            sb.AddSimpleField(F_USER_MACH, typeof(string));

            return sb.Finish();
        }

        // ══════════════════════════════════════════════════════════
        //  EXTENSIBLE STORAGE — Read / Write
        // ══════════════════════════════════════════════════════════

        /// <summary>
        /// Reads traceability data from a Family element's Extensible Storage.
        /// Returns null if no data has been stamped yet.
        /// Does NOT require a Transaction (read-only).
        /// </summary>
        public static TraceData ReadTraceData(Family family)
        {
            try
            {
                Schema schema = Schema.Lookup(SchemaGuid);
                if (schema == null) return null;

                Entity entity = family.GetEntity(schema);
                if (!entity.IsValid()) return null;

                return new TraceData
                {
                    AdcPath = entity.Get<string>(F_ADC_PATH) ?? "",
                    Sha256Hash = entity.Get<string>(F_SHA256) ?? "",
                    LoadTimestamp = entity.Get<string>(F_TIMESTAMP) ?? "",
                    UserMachineId = entity.Get<string>(F_USER_MACH) ?? ""
                };
            }
            catch (Exception) { return null; }
        }

        /// <summary>
        /// Writes traceability data to a Family element's Extensible Storage.
        /// MUST be called inside an open <see cref="Transaction"/>.
        /// </summary>
        public static void WriteTraceData(Family family, TraceData data)
        {
            Schema schema = GetOrCreateSchema();
            Entity entity = new Entity(schema);
            entity.Set(F_ADC_PATH, data.AdcPath ?? "");
            entity.Set(F_SHA256, data.Sha256Hash ?? "");
            entity.Set(F_TIMESTAMP, data.LoadTimestamp ?? "");
            entity.Set(F_USER_MACH, data.UserMachineId ?? "");
            family.SetEntity(entity);
        }

        // ══════════════════════════════════════════════════════════
        //  VISIBLE PARAMETER INJECTION
        // ══════════════════════════════════════════════════════════

        /// <summary>
        /// Opens the family document, adds or updates the "ADC_Library_Path"
        /// Type Parameter, then reloads the family into the project.
        /// 
        /// This is a HEAVYWEIGHT operation (~1-3 sec per family).
        /// It triggers a new FamilyLoaded event — callers must guard
        /// against recursion.
        /// 
        /// MUST be called from a valid Revit API context (main thread).
        /// Does NOT require an outer Transaction — it manages its own.
        /// </summary>
        /// <returns>True on success.</returns>
        public static bool InjectVisibleParameter(
            Document doc, Family family, string adcPath)
        {
            if (family == null || family.IsInPlace) return false;

            Document famDoc = null;
            try
            {
                famDoc = doc.EditFamily(family);
                if (famDoc == null) return false;

                FamilyManager mgr = famDoc.FamilyManager;

                // Check if parameter already exists
                FamilyParameter existing = null;
                foreach (FamilyParameter fp in mgr.GetParameters())
                {
                    if (fp.Definition.Name == PARAM_ADC_PATH)
                    { existing = fp; break; }
                }

                using (Transaction tx = new Transaction(famDoc, "HMV - Inject ADC Path"))
                {
                    tx.Start();
                    try
                    {
                        if (existing == null)
                        {
                            // Add as a Type Parameter under Identity Data group.
                            // SpecTypeId.String.Text = plain text data type.
                            existing = mgr.AddParameter(
                                PARAM_ADC_PATH,
                                GroupTypeId.IdentityData,
                                SpecTypeId.String.Text,
                                false);  // false = Type parameter (not instance)
                        }

                        // Ensure we have at least one type to set values on
                        if (mgr.CurrentType == null && mgr.Types.Size > 0)
                        {
                            var enumerator = mgr.Types.GetEnumerator();
                            if (enumerator.MoveNext())
                                mgr.CurrentType = enumerator.Current as FamilyType;
                        }

                        // Set the value on every FamilyType
                        foreach (FamilyType ft in mgr.Types)
                        {
                            mgr.CurrentType = ft;
                            mgr.Set(existing, adcPath ?? "");
                        }

                        tx.Commit();
                    }
                    catch
                    {
                        if (tx.HasStarted()) tx.RollBack();
                        return false;
                    }
                }

                // Reload the modified family back into the project
                famDoc.LoadFamily(doc, new TraceabilityLoadOptions());
                famDoc.Close(false);
                famDoc = null;
                return true;
            }
            catch (Exception)
            {
                // Close family doc without saving if something went wrong
                try { famDoc?.Close(false); } catch { }
                return false;
            }
        }

        // ══════════════════════════════════════════════════════════
        //  HELPERS
        // ══════════════════════════════════════════════════════════

        /// <summary>Builds a user+machine identifier for the trace stamp.</summary>
        public static string GetUserMachineId()
        {
            return $"{Environment.UserName}@{Environment.MachineName}";
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  IFamilyLoadOptions for silent reload during parameter injection
    // ═══════════════════════════════════════════════════════════════

    /// <summary>
    /// Always overwrites the existing family and its parameter values
    /// when reloading after parameter injection.
    /// </summary>
    internal class TraceabilityLoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse,
            out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;  // continue loading
        }

        public bool OnSharedFamilyFound(Family sharedFamily,
            bool familyInUse, out FamilySource source,
            out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }
}