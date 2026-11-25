// Compiled by Windows PowerShell 5.1. Read the copied OLE object, never the
// current Outlook selection (which may already have changed).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows.Forms;
using ComDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace OutlookLocalClipboard
{
    [ComImport, Guid("0000000b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IStorage
    {
        void CreateStream([MarshalAs(UnmanagedType.LPWStr)] string name, uint mode, uint r1, uint r2, out IStream stream);
        void OpenStream([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr reserved, uint mode, uint r2, out IStream stream);
        void CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string name, uint mode, uint r1, uint r2, out IStorage storage);
        void OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string name, IStorage priority, uint mode, IntPtr exclude, uint reserved, out IStorage storage);
        void CopyTo(uint count, IntPtr excludedIids, IntPtr excludedNames, IStorage destination);
        void MoveElementTo([MarshalAs(UnmanagedType.LPWStr)] string name, IStorage destination, [MarshalAs(UnmanagedType.LPWStr)] string newName, uint flags);
        void Commit(uint flags);
        void Revert();
        void EnumElements(uint r1, IntPtr r2, uint r3, out IntPtr enumerator);
        void DestroyElement([MarshalAs(UnmanagedType.LPWStr)] string name);
        void RenameElement([MarshalAs(UnmanagedType.LPWStr)] string oldName, [MarshalAs(UnmanagedType.LPWStr)] string newName);
        void SetElementTimes([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr created, IntPtr accessed, IntPtr modified);
        void SetClass(ref Guid clsid);
        void SetStateBits(uint bits, uint mask);
        void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG stat, uint flags);
    }

    public static class Native
    {
        [DllImport("ole32.dll")] public static extern int OleInitialize(IntPtr reserved);
        [DllImport("ole32.dll")] public static extern void OleUninitialize();
        [DllImport("ole32.dll")] public static extern int OleGetClipboard(out ComDataObject data);
        [DllImport("ole32.dll")] public static extern void ReleaseStgMedium(ref STGMEDIUM medium);
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)] public static extern int StgCreateDocfile(string path, uint mode, uint reserved, out IStorage storage);
        [DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber();
        [DllImport("user32.dll")] public static extern IntPtr GetClipboardOwner();
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern uint RegisterClipboardFormat(string name);
        [DllImport("user32.dll")] public static extern bool IsClipboardFormatAvailable(uint format);
        [DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr window);
        [DllImport("user32.dll")] public static extern bool CloseClipboard();
        [DllImport("user32.dll")] public static extern bool EmptyClipboard();
        [DllImport("user32.dll")] public static extern IntPtr SetClipboardData(uint format, IntPtr data);
        [DllImport("kernel32.dll")] public static extern IntPtr GlobalAlloc(uint flags, UIntPtr size);
        [DllImport("kernel32.dll")] public static extern IntPtr GlobalLock(IntPtr handle);
        [DllImport("kernel32.dll")] public static extern bool GlobalUnlock(IntPtr handle);
        [DllImport("kernel32.dll")] public static extern UIntPtr GlobalSize(IntPtr handle);
        [DllImport("kernel32.dll")] public static extern IntPtr GlobalFree(IntPtr handle);
    }

    public sealed class Bridge : IDisposable
    {
        public const long MaxFileBytes = 256L * 1024 * 1024;
        readonly string exportRoot;
        readonly NativeWindow window = new NativeWindow();
        uint seen;
        uint pendingSequence;
        string[] pending;
        int attempts;
        readonly uint descriptorW = Native.RegisterClipboardFormat("FileGroupDescriptorW");
        readonly uint descriptorA = Native.RegisterClipboardFormat("FileGroupDescriptor");
        readonly uint contents = Native.RegisterClipboardFormat("FileContents");

        public Bridge(string directory)
        {
            Marshal.ThrowExceptionForHR(Native.OleInitialize(IntPtr.Zero));
            exportRoot = Path.GetFullPath(directory);
            window.CreateHandle(new CreateParams { Caption = "Outlook Local Clipboard Bridge", Parent = new IntPtr(-3) });
            // Do not export whatever was on the clipboard before startup.
            seen = Native.GetClipboardSequenceNumber();
        }

        public string[] Poll()
        {
            Application.DoEvents();
            uint sequence = Native.GetClipboardSequenceNumber();
            if (sequence == seen) return null;
            if (pendingSequence != sequence) { pending = null; attempts = 0; pendingSequence = sequence; }
            try
            {
                if (pending == null)
                {
                    uint pid;
                    Native.GetWindowThreadProcessId(Native.GetClipboardOwner(), out pid);
                    bool outlook = false;
                    try { using (Process p = Process.GetProcessById((int)pid)) outlook = String.Equals(p.ProcessName, "OUTLOOK", StringComparison.OrdinalIgnoreCase); }
                    catch (ArgumentException) { }
                    if (!outlook || Native.IsClipboardFormatAvailable(15)) { seen = sequence; return null; }
                    bool unicode = Native.IsClipboardFormatAvailable(descriptorW);
                    uint descriptor = unicode ? descriptorW : descriptorA;
                    if (!Native.IsClipboardFormatAvailable(descriptor) || !Native.IsClipboardFormatAvailable(contents)) { seen = sequence; return null; }
                    ComDataObject data = null;
                    try
                    {
                        Marshal.ThrowExceptionForHR(Native.OleGetClipboard(out data));
                        if (sequence != Native.GetClipboardSequenceNumber()) return null;
                        string[] names = ReadNames(data, descriptor, unicode);
                        if (names == null) { seen = sequence; return null; }
                        pending = Export(data, names, exportRoot, contents);
                    }
                    finally { if (data != null && Marshal.IsComObject(data)) Marshal.ReleaseComObject(data); }
                }
                // OpenClipboard + sequence comparison makes replacement atomic with
                // respect to another application copying new text/files meanwhile.
                if (sequence != Native.GetClipboardSequenceNumber()) { pending = null; return null; }
                uint publishedSequence;
                if (!PublishFiles(pending, sequence, window.Handle, out publishedSequence))
                {
                    if (++attempts >= 20) throw new IOException("Clipboard remained busy; exported files are retained in the export folder. Copy the mail again.");
                    return null;
                }
                string[] result = pending;
                pending = null;
                seen = publishedSequence;
                return result;
            }
            catch (COMException error)
            {
                // Outlook may briefly reject OLE calls while finishing Copy.
                if ((error.ErrorCode == unchecked((int)0x800401D0) || error.ErrorCode == unchecked((int)0x80010001) || error.ErrorCode == unchecked((int)0x8001010A)) && ++attempts < 20) return null;
                seen = sequence;
                pending = null;
                throw;
            }
            catch
            {
                seen = sequence;
                pending = null;
                throw;
            }
        }

        static FORMATETC Format(uint id, int index, TYMED medium)
        {
            return new FORMATETC { cfFormat = unchecked((short)id), dwAspect = DVASPECT.DVASPECT_CONTENT, lindex = index, ptd = IntPtr.Zero, tymed = medium };
        }

        static byte[] ReadGlobal(IntPtr handle, long maximum)
        {
            ulong size = Native.GlobalSize(handle).ToUInt64();
            if (size == 0 || size > (ulong)maximum) throw new InvalidDataException("Invalid or oversized clipboard data.");
            IntPtr pointer = Native.GlobalLock(handle);
            if (pointer == IntPtr.Zero) throw new IOException("Could not lock clipboard data.");
            try { byte[] bytes = new byte[(int)size]; Marshal.Copy(pointer, bytes, 0, bytes.Length); return bytes; }
            finally { Native.GlobalUnlock(handle); }
        }

        static string[] ReadNames(ComDataObject data, uint id, bool unicode)
        {
            FORMATETC format = Format(id, -1, TYMED.TYMED_HGLOBAL);
            STGMEDIUM medium;
            data.GetData(ref format, out medium);
            try
            {
                if (medium.tymed != TYMED.TYMED_HGLOBAL) throw new InvalidDataException("Unsupported file descriptor storage.");
                return ParseNames(ReadGlobal(medium.unionmember, 1024 * 1024), unicode);
            }
            finally { Native.ReleaseStgMedium(ref medium); }
        }

        public static string[] ParseNames(byte[] bytes, bool unicode)
        {
            if (bytes.Length < 4) throw new InvalidDataException("Truncated file descriptor.");
            uint count = BitConverter.ToUInt32(bytes, 0);
            int stride = unicode ? 592 : 332;
            if (count == 0 || count > 20 || bytes.Length < 4L + count * stride) throw new InvalidDataException("Expected 1 to 20 complete mail descriptors.");
            string[] names = new string[count];
            for (int i = 0; i < count; i++)
            {
                string name = (unicode ? Encoding.Unicode : Encoding.Default).GetString(bytes, 4 + i * stride + 72, unicode ? 520 : 260).Split('\0')[0];
                if (!String.Equals(Path.GetExtension(name), ".msg", StringComparison.OrdinalIgnoreCase)) return null;
                names[i] = name;
            }
            return names;
        }

        public static string SafeName(string name, int index)
        {
            string stem = Path.GetFileNameWithoutExtension(name.Replace('/', '\\'));
            foreach (char c in Path.GetInvalidFileNameChars()) stem = stem.Replace(c, '_');
            stem = stem.Trim().TrimEnd('.');
            if (stem.Length > 80) stem = stem.Substring(0, 80);
            if (stem.Length == 0) stem = "mail";
            // A numeric prefix also neutralizes reserved names such as CON.
            return (index + 1).ToString("D2") + "-" + stem + ".msg";
        }

        public static string[] Export(ComDataObject data, string[] names, string root, uint contentsFormat)
        {
            string directory = Path.Combine(root, DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 12));
            Directory.CreateDirectory(directory);
            List<string> paths = new List<string>();
            try
            {
                long total = 0;
                for (int i = 0; i < names.Length; i++)
                {
                    string file = Path.Combine(directory, SafeName(names[i], i));
                    paths.Add(file);
                    FORMATETC format = Format(contentsFormat, i, TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL);
                    STGMEDIUM medium;
                    data.GetData(ref format, out medium);
                    try { SaveMedium(ref medium, file); }
                    finally { Native.ReleaseStgMedium(ref medium); }
                    ValidateMsg(file);
                    total += new FileInfo(file).Length;
                    if (total > MaxFileBytes) throw new InvalidDataException("Copied mail exceeds the 256 MiB batch limit.");
                }
                return paths.ToArray();
            }
            catch
            {
                // Only remove exact files created by this export, never recurse.
                foreach (string file in paths) { try { File.Delete(file); } catch { } }
                try { Directory.Delete(directory); } catch { }
                throw;
            }
        }

        public static void SaveMedium(ref STGMEDIUM medium, string file)
        {
            if (medium.tymed == TYMED.TYMED_HGLOBAL)
            {
                byte[] bytes = ReadGlobal(medium.unionmember, MaxFileBytes);
                using (FileStream output = new FileStream(file, FileMode.CreateNew)) output.Write(bytes, 0, bytes.Length);
            }
            else if (medium.tymed == TYMED.TYMED_ISTREAM)
            {
                IStream stream = (IStream)Marshal.GetObjectForIUnknown(medium.unionmember);
                IntPtr read = Marshal.AllocHGlobal(4);
                try
                {
                    stream.Seek(0, 0, IntPtr.Zero);
                    using (FileStream output = new FileStream(file, FileMode.CreateNew))
                    {
                        byte[] buffer = new byte[65536];
                        while (true)
                        {
                            Marshal.WriteInt32(read, 0);
                            stream.Read(buffer, buffer.Length, read);
                            int length = Marshal.ReadInt32(read);
                            if (length == 0) break;
                            if (length < 0 || length > buffer.Length || output.Length + length > MaxFileBytes) throw new InvalidDataException("Invalid or oversized mail stream.");
                            output.Write(buffer, 0, length);
                        }
                    }
                }
                finally { Marshal.FreeHGlobal(read); if (Marshal.IsComObject(stream)) Marshal.ReleaseComObject(stream); }
            }
            else if (medium.tymed == TYMED.TYMED_ISTORAGE)
            {
                IStorage source = (IStorage)Marshal.GetObjectForIUnknown(medium.unionmember);
                IStorage destination = null;
                try
                {
                    // STGM_READWRITE | STGM_SHARE_EXCLUSIVE; fail if already present.
                    Marshal.ThrowExceptionForHR(Native.StgCreateDocfile(file, 0x12, 0, out destination));
                    source.CopyTo(0, IntPtr.Zero, IntPtr.Zero, destination);
                    System.Runtime.InteropServices.ComTypes.STATSTG stat;
                    source.Stat(out stat, 1);
                    destination.SetClass(ref stat.clsid);
                    destination.Commit(0);
                }
                finally
                {
                    if (destination != null) Marshal.ReleaseComObject(destination);
                    if (Marshal.IsComObject(source)) Marshal.ReleaseComObject(source);
                }
            }
            else throw new InvalidDataException("Unsupported mail clipboard storage.");
        }

        public static void ValidateMsg(string file)
        {
            byte[] expected = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
            using (FileStream input = File.OpenRead(file))
            {
                if (input.Length < 512 || input.Length > MaxFileBytes) throw new InvalidDataException("Invalid or oversized MSG compound file.");
                foreach (byte value in expected) if (input.ReadByte() != value) throw new InvalidDataException("Copied object is not an MSG compound file.");
            }
        }

        public static byte[] FileDropBytes(string[] paths)
        {
            byte[] names = Encoding.Unicode.GetBytes(String.Join("\0", paths) + "\0\0");
            byte[] bytes = new byte[20 + names.Length];
            BitConverter.GetBytes(20).CopyTo(bytes, 0); // DROPFILES.pFiles
            BitConverter.GetBytes(1).CopyTo(bytes, 16); // DROPFILES.fWide
            names.CopyTo(bytes, 20);
            return bytes;
        }

        static IntPtr Allocate(byte[] bytes)
        {
            IntPtr memory = Native.GlobalAlloc(0x42, new UIntPtr((uint)bytes.Length));
            if (memory == IntPtr.Zero) throw new OutOfMemoryException();
            IntPtr pointer = Native.GlobalLock(memory);
            if (pointer == IntPtr.Zero) { Native.GlobalFree(memory); throw new IOException("Could not allocate file clipboard."); }
            try { Marshal.Copy(bytes, 0, pointer, bytes.Length); }
            finally { Native.GlobalUnlock(memory); }
            return memory;
        }

        public static bool PublishFiles(string[] paths, uint sequence, IntPtr owner, out uint publishedSequence)
        {
            publishedSequence = sequence;
            IntPtr memory = Allocate(FileDropBytes(paths));
            bool opened = false;
            try
            {
                opened = Native.OpenClipboard(owner);
                if (!opened || Native.GetClipboardSequenceNumber() != sequence) return false;
                if (!Native.EmptyClipboard()) throw new IOException("Could not prepare file clipboard.");
                if (Native.SetClipboardData(15, memory) == IntPtr.Zero) throw new IOException("Could not publish file clipboard; exported files are retained.");
                memory = IntPtr.Zero; // Windows owns the memory after success.
                publishedSequence = Native.GetClipboardSequenceNumber();
                return true;
            }
            finally { if (opened) Native.CloseClipboard(); if (memory != IntPtr.Zero) Native.GlobalFree(memory); }
        }

        public void Dispose() { window.DestroyHandle(); Native.OleUninitialize(); }
    }
}
