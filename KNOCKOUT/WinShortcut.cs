using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace KNOCKOUT
{
    /// <summary>
    /// Represents partial properties of a Windows shortcut file.
    /// </summary>
    public class WinShortcut
    {
        private string _hotKey;

        /// <summary>
        /// Initialize an instance of this class using the path of shortcut file.
        /// </summary>
        /// <param name="path">The path of the shortcut file</param>
        public WinShortcut(string path)
        {
            using (var istream = File.OpenRead(path))
            {
                try
                {
                    this.Parse(istream);
                }
                catch (Exception ex)
                {
                    throw new Exception("Failed to parse this file as a Windows shortcut", ex);
                }
            }
        }

        /// <summary>
        /// The real path of target this shortcut refers to.
        /// </summary>
        public string TargetPath { get; private set; }

        /// <summary>
        /// Whether the target this shortcut refers to is a directory.
        /// </summary>
        public bool IsDirectory { get; private set; }

        /// <summary>
        /// Hotkey of this shortcut.
        /// </summary>
        public string HotKey
        {
            get { return this._hotKey ?? ""; }
            private set { _hotKey = value; }
        }

        private void Parse(Stream istream)
        {
            var linkFlags = this.ParseHeader(istream);
            if ((linkFlags & ShortcutConstants.LinkFlags.HasLinkTargetIdList) == ShortcutConstants.LinkFlags.HasLinkTargetIdList)
            {
                this.ParseTargetIDList(istream);
            }
            if ((linkFlags & ShortcutConstants.LinkFlags.HasLinkInfo) == ShortcutConstants.LinkFlags.HasLinkInfo)
            {
                this.ParseLinkInfo(istream);
            }
        }

        /// <summary>
        /// Parse the header.
        /// </summary>
        /// <param name="stream"></param>
        /// <returns>The flags that specify the presence of optional structures</returns>
        private int ParseHeader(Stream stream)
        {
            stream.Seek(20, SeekOrigin.Begin);//jump to the LinkFlags part of ShellLinkHeader
            var buffer = new byte[4];
            stream.Read(buffer, 0, buffer.Length);
            var linkFlags = BitConverter.ToInt32(buffer, 0);

            stream.Read(buffer, 0, buffer.Length);//read next 4 bytes, that is FileAttributes
            var fileAttrFlags = BitConverter.ToInt32(buffer, 0);
            IsDirectory = (fileAttrFlags & ShortcutConstants.FileAttributes.Directory) == ShortcutConstants.FileAttributes.Directory;

            stream.Seek(36, SeekOrigin.Current);//jump to the HotKey part
            stream.Read(buffer, 0, 2);

            var keys = new List<string>();
            var hotKeyLowByte = (ShortcutConstants.VirtualKeys)buffer[0];
            var hotKeyHighByte = (ShortcutConstants.VirtualKeys)buffer[1];
            if (hotKeyHighByte.HasFlag(ShortcutConstants.VirtualKeys.HOTKEYF_CONTROL))
                keys.Add("ctrl");
            if (hotKeyHighByte.HasFlag(ShortcutConstants.VirtualKeys.HOTKEYF_SHIFT))
                keys.Add("shift");
            if (hotKeyHighByte.HasFlag(ShortcutConstants.VirtualKeys.HOTKEYF_ALT))
                keys.Add("alt");
            if (Enum.IsDefined(typeof(ShortcutConstants.VirtualKeys), hotKeyLowByte))
                keys.Add(hotKeyLowByte.ToString());
            HotKey = String.Join("+", keys);

            return linkFlags;
        }

        /// <summary>
        /// Parse the TargetIDList part.
        /// </summary>
        /// <param name="stream"></param>
        private void ParseTargetIDList(Stream stream)
        {
            stream.Seek(76, SeekOrigin.Begin);//jump to the LinkTargetIDList part
            var buffer = new byte[2];
            stream.Read(buffer, 0, buffer.Length);
            var size = BitConverter.ToInt16(buffer, 0);
            //the TargetIDList part isn't used currently, so just move the cursor forward
            stream.Seek(size, SeekOrigin.Current);
        }

        /// <summary>
        /// Parse the LinkInfo part.
        /// </summary>
        /// <param name="stream"></param>
        private void ParseLinkInfo(Stream stream)
        {
            var start = stream.Position;//save the start position of LinkInfo
            stream.Seek(8, SeekOrigin.Current);//jump to the LinkInfoFlags part
            var buffer = new byte[4];
            stream.Read(buffer, 0, buffer.Length);
            var lnkInfoFlags = BitConverter.ToInt32(buffer, 0);
            if ((lnkInfoFlags & ShortcutConstants.LinkInfoFlags.VolumeIDAndLocalBasePath) == ShortcutConstants.LinkInfoFlags.VolumeIDAndLocalBasePath)
            {
                stream.Seek(4, SeekOrigin.Current);
                stream.Read(buffer, 0, buffer.Length);
                var localBasePathOffset = BitConverter.ToInt32(buffer, 0);
                var basePathOffset = start + localBasePathOffset;
                stream.Seek(basePathOffset, SeekOrigin.Begin);

                using (var ms = new MemoryStream())
                {
                    var b = 0;
                    //get raw bytes of LocalBasePath
                    while ((b = stream.ReadByte()) > 0)
                        ms.WriteByte((byte)b);

                    TargetPath = Encoding.Default.GetString(ms.ToArray());
                }
            }
        }
    }
}

