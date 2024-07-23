using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace KNOCKOUT
{
    internal class RegistryHelper
    {
        // P/Invoke constants
        private const int READ_CONTROL = 0x00020000;
        private const int KEY_QUERY_VALUE = 0x0001;
        private const int KEY_ENUMERATE_SUB_KEYS = 0x0008;
        private const int KEY_READ = READ_CONTROL | KEY_QUERY_VALUE | KEY_ENUMERATE_SUB_KEYS;
        const int KEY_WOW64_32KEY = 0x0200;
        const int KEY_WOW64_64KEY = 0x0100;
        const int ERROR_SUCCESS = 0;
        const int ERROR_NO_MORE_ITEMS = 259;
        // Registry hive definition
        public static UIntPtr HKEY_LOCAL_MACHINE = new UIntPtr(0x80000002);
        public static UIntPtr HKEY_CURRENT_USER = new UIntPtr(0x80000001);

        // P/Invoke structures
        [StructLayout(LayoutKind.Sequential)]
        public struct FILETIME
        {
            public uint dwLowDateTime;
            public uint dwHighDateTime;
        }

        // P/Invoke functions
        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern int RegOpenKeyEx(
            UIntPtr hKey,
            string subKey,
            uint options,
            int samDesired,
            out IntPtr phkResult);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern int RegCloseKey(IntPtr hKey);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern int RegEnumKeyEx(
        IntPtr hKey,
        uint dwIndex,
        StringBuilder lpName,
        ref uint lpcbName,
        IntPtr lpReserved,
        StringBuilder lpClass,
        ref uint lpcbClass,
        out long lpftLastWriteTime);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int RegEnumValue(
            IntPtr hKey,
            uint index,
            StringBuilder lpValueName,
            ref uint lpcbValueName,
            IntPtr lpReserved,
            out uint lpType,
            IntPtr lpData,
            ref uint lpcbData
        );

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
            public static extern int RegQueryValueEx(
            IntPtr hKey,
            string lpValueName,
            IntPtr lpReserved,
            ref uint lpType,
            IntPtr lpData,
            ref uint lpcbData
        );

        const uint REG_BINARY = 3;

        // Opens the specified registry hive and path.
        public IntPtr OpenRegistryKey(UIntPtr hive, string path)
        {
            IntPtr hRootKey;
            int result = RegOpenKeyEx(hive, path, 0, KEY_READ, out hRootKey);
            if (result != 0)
            {
                throw new Exception("Failed to open registry key.");
            }
            return hRootKey;
        }

        // Enumerates the subkeys of the given key handle and returns them as a list.
        public List<string> EnumerateSubKeys(IntPtr hKey)
        {
            List<string> subKeys = new List<string>();
            uint index = 0;
            uint nameLength = 1024;
            StringBuilder nameBuffer = new StringBuilder((int)nameLength);
            long lastWriteTime;
            int result;

            while ((result = RegEnumKeyEx(hKey, index, nameBuffer, ref nameLength, IntPtr.Zero, null, ref nameLength, out lastWriteTime)) == ERROR_SUCCESS)
            {
                subKeys.Add(nameBuffer.ToString());
                nameBuffer.Length = 0; // Reset the buffer for the next subkey
                nameLength = 1024; // Reset the length for the next subkey
                index++;
            }

            if (result != 0 && result != 259) // ERROR_NO_MORE_ITEMS
            {
                throw new Exception("An error occurred while enumerating subkeys.");
            }

            return subKeys;
        }

        // Method to recursively traverse the registry from the given path and find specific subkeys.
        public List<string> FindMRUKeys(IntPtr hKey, string currentPath)
        {
            List<string> mruPaths = new List<string>();
            uint index = 0;
            uint nameLength = 1024;
            StringBuilder nameBuffer = new StringBuilder((int)nameLength);
            long lastWriteTime;
            int result;

            while ((result = RegEnumKeyEx(hKey, index, nameBuffer, ref nameLength, IntPtr.Zero, null, ref nameLength, out lastWriteTime)) == ERROR_SUCCESS)
            {
                string subKeyName = nameBuffer.ToString();
                string fullPath = $"{currentPath}\\{subKeyName}";

                // Check if the subkey name is "File MRU" or "Place MRU"
                if (subKeyName.Equals("File MRU", StringComparison.OrdinalIgnoreCase) ||
                    subKeyName.Equals("Place MRU", StringComparison.OrdinalIgnoreCase))
                {
                    mruPaths.Add(fullPath);
                }

                // Open the subkey and continue the recursive search
                IntPtr hSubKey;
                result = RegOpenKeyEx((UIntPtr)hKey.ToInt64(), subKeyName, 0, KEY_READ, out hSubKey);
                if (result == ERROR_SUCCESS)
                {
                    mruPaths.AddRange(FindMRUKeys(hSubKey, fullPath)); // Recursive call
                    RegCloseKey(hSubKey);
                }

                nameBuffer.Length = 0; // Reset the buffer for the next subkey
                nameLength = 1024; // Reset the length for the next subkey
                index++;
            }

            return mruPaths;
        }

        // Function to retrieve all values and their data from a registry key
        public Dictionary<string, string> RetrieveRegistryValues(IntPtr hKey)
        {
            Dictionary<string, string> valuesData = new Dictionary<string, string>();

            uint index = 0;
            uint dataSize = 1024;
            StringBuilder valueNameBuffer = new StringBuilder((int)dataSize);
            uint type = 0;

            while (RegEnumValue(hKey, index, valueNameBuffer, ref dataSize, IntPtr.Zero, out type, IntPtr.Zero, ref dataSize) == ERROR_SUCCESS)
            {
                IntPtr dataPtr = Marshal.AllocHGlobal((int)dataSize);
                try
                {
                    if (RegQueryValueEx(hKey, valueNameBuffer.ToString(), IntPtr.Zero, ref type, dataPtr, ref dataSize) == ERROR_SUCCESS)
                    {
                        string valueData = GetValueData(dataPtr, type, dataSize);
                        valuesData.Add(valueNameBuffer.ToString(), valueData);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(dataPtr);
                }

                dataSize = 1024;
                index++;
            }

            return valuesData;
        }

        // Function to retrieve all values and their data from a registry key
        public Dictionary<string, byte[]> RetrieveRegistryValuesRaw(IntPtr hKey)
        {
            Dictionary<string, byte[]> valuesData = new Dictionary<string, byte[]>();

            uint index = 0;
            uint dataSize = 1024;
            StringBuilder valueNameBuffer = new StringBuilder((int)dataSize);
            uint type = 0;

            while (RegEnumValue(hKey, index, valueNameBuffer, ref dataSize, IntPtr.Zero, out type, IntPtr.Zero, ref dataSize) == ERROR_SUCCESS)
            {
                IntPtr dataPtr = Marshal.AllocHGlobal((int)dataSize);
                try
                {
                    byte[] data = GetValueDataRaw(hKey, valueNameBuffer.ToString());
                    valuesData.Add(valueNameBuffer.ToString(), data);
                }
                finally
                {
                    Marshal.FreeHGlobal(dataPtr);
                }

                dataSize = 1024;
                index++;
            }

            return valuesData;
        }

        static byte[] GetValueDataRaw(IntPtr hKey, string valueName)
        {
            uint type = 0;
            uint dataSize = 0;

            // Query the value to get the size of the data
            int result = RegQueryValueEx(hKey, valueName, IntPtr.Zero, ref type, IntPtr.Zero, ref dataSize);
            if (result != ERROR_SUCCESS)
            {
                throw new Exception("Error querying registry value size");
            }

            // Allocate a buffer for the data
            //byte[] data = new byte[dataSize];
            IntPtr dataPtr = Marshal.AllocHGlobal((int)dataSize);

            // Query the value again to get the actual data
            result = RegQueryValueEx(hKey, valueName, IntPtr.Zero, ref type, dataPtr, ref dataSize);
            if (result != ERROR_SUCCESS)
            {
                throw new Exception("Error querying registry value data");
            }

            // Ensure the type is REG_BINARY
            if (type != REG_BINARY)
            {
                throw new Exception("Registry value is not of type REG_BINARY");
            }

            byte[] binaryData = GetBytesFromIntPtr(dataPtr, dataSize);

            return binaryData;
        }

        public Dictionary<string, string> RetrieveRegistryValuesBinary(IntPtr hKey)
        {
            Dictionary<string, string> valuesData = new Dictionary<string, string>();

            uint index = 0;
            uint dataSize = 1024;
            StringBuilder valueNameBuffer = new StringBuilder((int)dataSize);
            uint type = 0;

            while (RegEnumValue(hKey, index, valueNameBuffer, ref dataSize, IntPtr.Zero, out type, IntPtr.Zero, ref dataSize) == 0)
            {
                IntPtr dataPtr = Marshal.AllocHGlobal((int)dataSize);
                try
                {
                    if (RegQueryValueEx(hKey, valueNameBuffer.ToString(), IntPtr.Zero, ref type, dataPtr, ref dataSize) == 0)
                    {
                        string valueData = GetValueDataBinary(dataPtr, type, dataSize);
                        valuesData.Add(valueNameBuffer.ToString(), valueData);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(dataPtr);
                }

                dataSize = 1024;
                index++;
            }

            return valuesData;
        }

        // Function to retrieve value data based on its type
        private static string GetValueData(IntPtr dataPtr, uint type, uint dataSize)
        {
            switch (type)
            {
                case 1: // REG_SZ
                    return Marshal.PtrToStringUni(dataPtr);
                case 2: // REG_EXPAND_SZ
                    return Marshal.PtrToStringUni(dataPtr);
                case 4: // REG_DWORD
                    return BitConverter.ToUInt32(GetBytesFromIntPtr(dataPtr, dataSize), 0).ToString();

                case 7: // REG_MULTI_SZ
                    return Marshal.PtrToStringAuto(dataPtr);

                case 11: // REG_QWORD
                    return BitConverter.ToUInt64(GetBytesFromIntPtr(dataPtr, dataSize), 0).ToString();

                default:
                    return "Unsupported data type";
            }
        }

        // Special function to treat REG_BINARY strings of the RecentDocs
        private static string GetValueDataBinary(IntPtr dataPtr, uint type, uint dataSize)
        {
            switch (type)
            {
                case 3: // REG_BINARY
                    byte[] binaryData = GetBytesFromIntPtr(dataPtr, dataSize);
                    string dataStr = Encoding.Unicode.GetString(binaryData);
                    string file = dataStr.Split('\0')[0];
                    return file;

                default:
                    return "Unsupported data type";
            }
        }

        // Helper function to get byte array from IntPtr
        private static byte[] GetBytesFromIntPtr(IntPtr dataPtr, uint dataSize)
        {
            byte[] bytes = new byte[dataSize];
            Marshal.Copy(dataPtr, bytes, 0, (int)dataSize);
            return bytes;
        }

        // Closes the registry key.
        public void CloseRegistryKey(IntPtr hKey)
        {
            RegCloseKey(hKey);
        }
    }
}
