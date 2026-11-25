using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace OutlookLocalClipboard
{
    public static class Tests
    {
        [DllImport("ole32.dll", CharSet = CharSet.Unicode)] static extern int StgOpenStorage(string path, IStorage priority, uint mode, IntPtr exclude, uint reserved, out IStorage storage);
        [DllImport("ole32.dll")] static extern int CreateStreamOnHGlobal(IntPtr memory, bool deleteOnRelease, out IStream stream);
        static void Assert(bool result, string message) { if (!result) throw new Exception(message); }
        static void Fails(Action action)
        {
            bool failed = false;
            try { action(); } catch (IOException) { failed = true; } catch (InvalidDataException) { failed = true; }
            Assert(failed, "Expected invalid input to fail");
        }
        static void Equal(byte[] a, byte[] b)
        {
            Assert(a.Length == b.Length, "Payload length changed");
            for (int i = 0; i < a.Length; i++) Assert(a[i] == b[i], "Payload bytes changed");
        }

        sealed class MailData : IDataObject
        {
            readonly byte[] bytes;
            readonly bool failSecond;
            public MailData(byte[] value, bool fail) { bytes = value; failSecond = fail; }
            public void GetData(ref FORMATETC format, out STGMEDIUM medium)
            {
                if (failSecond && format.lindex == 1) throw new IOException("Synthetic second mail failure");
                IntPtr memory = Native.GlobalAlloc(0x42, new UIntPtr((uint)bytes.Length));
                IntPtr pointer = Native.GlobalLock(memory);
                Marshal.Copy(bytes, 0, pointer, bytes.Length);
                Native.GlobalUnlock(memory);
                medium = new STGMEDIUM { tymed = TYMED.TYMED_HGLOBAL, unionmember = memory };
            }
            public void GetDataHere(ref FORMATETC f, ref STGMEDIUM m) { throw new NotImplementedException(); }
            public int QueryGetData(ref FORMATETC f) { return 0; }
            public int GetCanonicalFormatEtc(ref FORMATETC input, out FORMATETC output) { output = input; return 1; }
            public void SetData(ref FORMATETC f, ref STGMEDIUM m, bool release) { throw new NotImplementedException(); }
            public IEnumFORMATETC EnumFormatEtc(DATADIR d) { throw new NotImplementedException(); }
            public int DAdvise(ref FORMATETC f, ADVF flags, IAdviseSink sink, out int connection) { connection = 0; return unchecked((int)0x80004001); }
            public void DUnadvise(int connection) { throw new NotImplementedException(); }
            public int EnumDAdvise(out IEnumSTATDATA data) { data = null; return unchecked((int)0x80004001); }
        }

        public static void Run(string directory)
        {
            byte[] descriptor = new byte[4 + 2 * 592];
            BitConverter.GetBytes(2).CopyTo(descriptor, 0);
            string name = "\u90ae\u4ef6.msg";
            Encoding.Unicode.GetBytes(name).CopyTo(descriptor, 76);
            Encoding.Unicode.GetBytes(name).CopyTo(descriptor, 76 + 592);
            string[] names = Bridge.ParseNames(descriptor, true);
            Assert(names.Length == 2 && names[0] == name, "Unicode descriptor decoding");
            Assert(Bridge.SafeName(names[0], 0) != Bridge.SafeName(names[1], 1), "Duplicate filenames collide");
            Assert(Bridge.SafeName("../../CON.msg", 0) == "01-CON.msg", "Unsafe filename");
            Fails(delegate { Bridge.ParseNames(new byte[] { 1, 0, 0, 0 }, true); });
            descriptor[76] = descriptor[77] = 0;
            Assert(Bridge.ParseNames(descriptor, true) == null, "Non-MSG descriptor not ignored");
            BitConverter.GetBytes(21).CopyTo(descriptor, 0);
            Fails(delegate { Bridge.ParseNames(descriptor, true); });
            byte[] ansi = new byte[336];
            BitConverter.GetBytes(1).CopyTo(ansi, 0);
            Encoding.ASCII.GetBytes("mail.MSG").CopyTo(ansi, 76);
            Assert(Bridge.ParseNames(ansi, false)[0] == "mail.MSG", "ANSI descriptor decoding");

            string[] paths = { @"C:\test\mail.msg", @"C:\test\" + name };
            byte[] drop = Bridge.FileDropBytes(paths);
            Assert(BitConverter.ToInt32(drop, 0) == 20 && BitConverter.ToInt32(drop, 16) == 1, "Invalid DROPFILES header");
            Assert(Encoding.Unicode.GetString(drop, 20, drop.Length - 20) == String.Join("\0", paths) + "\0\0", "Invalid file list terminator");

            byte[] payload = Encoding.UTF8.GetBytes("Synthetic attachment payload");
            string sourcePath = Path.Combine(directory, "source.msg");
            string copyPath = Path.Combine(directory, "copy.msg");
            IStorage source;
            Marshal.ThrowExceptionForHR(Native.StgCreateDocfile(sourcePath, 0x12, 0, out source));
            try
            {
                IStorage attachment;
                source.CreateStorage("__attach_version1.0_#00000000", 0x12, 0, 0, out attachment);
                try
                {
                    IStream stream;
                    attachment.CreateStream("__substg1.0_37010102", 0x12, 0, 0, out stream);
                    try { stream.Write(payload, payload.Length, IntPtr.Zero); stream.Commit(0); }
                    finally { Marshal.ReleaseComObject(stream); }
                    attachment.Commit(0);
                }
                finally { Marshal.ReleaseComObject(attachment); }
                source.Commit(0);
                STGMEDIUM medium = new STGMEDIUM { tymed = TYMED.TYMED_ISTORAGE, unionmember = Marshal.GetIUnknownForObject(source) };
                try { Bridge.SaveMedium(ref medium, copyPath); }
                finally { Native.ReleaseStgMedium(ref medium); }
                Bridge.ValidateMsg(copyPath);
            }
            finally { Marshal.ReleaseComObject(source); }

            IStorage copied;
            Marshal.ThrowExceptionForHR(StgOpenStorage(copyPath, null, 0x10, IntPtr.Zero, 0, out copied));
            try
            {
                IStorage attachment;
                copied.OpenStorage("__attach_version1.0_#00000000", null, 0x10, IntPtr.Zero, 0, out attachment);
                try
                {
                    IStream stream;
                    attachment.OpenStream("__substg1.0_37010102", IntPtr.Zero, 0x10, 0, out stream);
                    try { byte[] bytes = new byte[payload.Length]; stream.Read(bytes, bytes.Length, IntPtr.Zero); Equal(bytes, payload); }
                    finally { Marshal.ReleaseComObject(stream); }
                }
                finally { Marshal.ReleaseComObject(attachment); }
            }
            finally { Marshal.ReleaseComObject(copied); }

            byte[] raw = File.ReadAllBytes(copyPath);
            string successRoot = Path.Combine(directory, "batch-success");
            string[] exported = Bridge.Export(new MailData(raw, false), new string[] { name, name }, successRoot, 123);
            Assert(exported.Length == 2 && exported[0] != exported[1], "Batch filename collision");
            Equal(File.ReadAllBytes(exported[0]), raw);
            Equal(File.ReadAllBytes(exported[1]), raw);
            string failedRoot = Path.Combine(directory, "batch-failure");
            Fails(delegate { Bridge.Export(new MailData(raw, true), new string[] { name, name }, failedRoot, 123); });
            Assert(Directory.GetFileSystemEntries(failedRoot).Length == 0, "Failed batch retained partial files");
            IntPtr memory = Native.GlobalAlloc(0x42, new UIntPtr((uint)raw.Length));
            Assert(memory != IntPtr.Zero, "GlobalAlloc failed");
            IntPtr pointer = Native.GlobalLock(memory);
            Marshal.Copy(raw, 0, pointer, raw.Length);
            Native.GlobalUnlock(memory);
            STGMEDIUM globalMedium = new STGMEDIUM { tymed = TYMED.TYMED_HGLOBAL, unionmember = memory };
            try
            {
                string globalPath = Path.Combine(directory, "global.msg");
                Bridge.SaveMedium(ref globalMedium, globalPath);
                Bridge.ValidateMsg(globalPath);
                Equal(File.ReadAllBytes(globalPath), raw);
                Fails(delegate { Bridge.SaveMedium(ref globalMedium, globalPath); });
            }
            finally { Native.ReleaseStgMedium(ref globalMedium); }

            IStream input;
            Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out input));
            try
            {
                input.Write(raw, raw.Length, IntPtr.Zero);
                STGMEDIUM medium = new STGMEDIUM { tymed = TYMED.TYMED_ISTREAM, unionmember = Marshal.GetIUnknownForObject(input) };
                try
                {
                    string file = Path.Combine(directory, "stream.msg");
                    Bridge.SaveMedium(ref medium, file);
                    Equal(File.ReadAllBytes(file), raw);
                }
                finally { Native.ReleaseStgMedium(ref medium); }
            }
            finally { Marshal.ReleaseComObject(input); }
            string invalid = Path.Combine(directory, "invalid.msg");
            File.WriteAllBytes(invalid, new byte[512]);
            Fails(delegate { Bridge.ValidateMsg(invalid); });
        }
    }
}
