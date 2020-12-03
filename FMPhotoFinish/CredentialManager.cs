/*
---
name: CredentialManager.cs
description: Store and retrieve credentials to/from Windows Credential Manager
url: 
version: 1.0
keywords: CodeBit
dateModified: 2020-12-03
license: https://opensource.org/licenses/BSD-3-Clause
# Metadata in MicroYaml format. See http://filemeta.org/CodeBit.html
...
*/

/*
=== BSD 3 Clause License ===
https://opensource.org/licenses/BSD-3-Clause

Copyright 2019 Brandt Redd

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors
may be used to endorse or promote products derived from this software without
specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.
*/

using System;
using System.Runtime.InteropServices;
using System.Text;

namespace FileMeta
{
    /// <summary>
    /// Store and retrieve secure credentials in the Windows Credential Manager
    /// </summary>
    static class CredentialManager
    {
        static readonly UTF8Encoding s_Utf8 = new UTF8Encoding(false);

        public static void Add(string targetName, string userName, string credential)
        {
            Add(targetName, userName, s_Utf8.GetBytes(credential));
        }

        public static void Add(string targetName, string userName, byte[] credential)
        {
            CREDENTIAL cred = new CREDENTIAL();
            try
            {
                // Fill in the credential structure
                cred.Type = CRED_TYPE.GENERIC;
                cred.TargetName = targetName;
                cred.CredentialBlobSize = (uint)credential.Length;
                cred.CredentialBlob = Marshal.AllocCoTaskMem(credential.Length);
                Marshal.Copy(credential, 0, cred.CredentialBlob, credential.Length);
                cred.Persist = CRED_PERSIST.ENTERPRISE;
                cred.UserName = userName;
                // All other values remain zeros

                // Write the credential
                CredWrite(cred, 0);
            }
            finally
            {
                // free memory
                if (cred.CredentialBlob != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(cred.CredentialBlob);
                    cred.CredentialBlob = IntPtr.Zero;
                }
            }
        }

        public static bool Retrieve(string name, out string userName, out string credential)
        {
            byte[] bytes;
            if (!Retrieve(name, out userName, out bytes))
            {
                credential = null;
                return false;
            }
            credential = s_Utf8.GetString(bytes);
            return true;
        }

        public static bool Retrieve(string name, out string username, out byte[] credential)
        {
            IntPtr bufPtr = IntPtr.Zero;
            try
            {
                if (!CredRead(name, CRED_TYPE.GENERIC, 0, out bufPtr))
                {
                    username = null;
                    credential = null;
                    return false;
                }
                CREDENTIAL cred = new CREDENTIAL();
                Marshal.PtrToStructure(bufPtr, cred);

                credential = new byte[cred.CredentialBlobSize];
                Marshal.Copy(cred.CredentialBlob, credential, 0, credential.Length);
                username = cred.UserName;
                return true;
            }
            finally
            {
                if (bufPtr != IntPtr.Zero)
                {
                    CredFree(bufPtr);
                    bufPtr = IntPtr.Zero;
                }
            }
        }

        public static bool Delete(string targetName)
        {
            return CredDelete(targetName, CRED_TYPE.GENERIC, 0);
        }

        public static string[] Enumerate(string filter)
        {
            UInt32 count = 0;
            IntPtr credListPtr = IntPtr.Zero;
            try
            {
                if (!CredEnumerate(filter, 0, out count, out credListPtr) || count == 0)
                {
                    return new string[0];
                }
                IntPtr[] credList = new IntPtr[count];
                Marshal.Copy(credListPtr, credList, 0, (int)count);
                string[] result = new string[count];
                int index = 0;
                foreach (IntPtr credPtr in credList)
                {
                    CREDENTIAL cred = new CREDENTIAL();
                    Marshal.PtrToStructure(credPtr, cred);
                    result[index] = cred.TargetName;
                    ++index;
                }
                return result;
            }
            finally
            {
                if (credListPtr != IntPtr.Zero)
                {
                    CredFree(credListPtr);
                    credListPtr = IntPtr.Zero;
                }
            }
        }

        #region PInvoke stuff

        [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredWrite([In] CREDENTIAL userCredential, [In] UInt32 flags);

        [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredRead([In] string targetName, [In] CRED_TYPE type, [In] int reservedFlag, [Out] out IntPtr credentialPtr);

        [DllImport("advapi32.dll", EntryPoint = "CredEnumerateW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredEnumerate([In] string filter, [In] UInt32 flags, [Out] out UInt32 Count, [Out] out IntPtr credentialsPtr);

        [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredDelete(string targetName, CRED_TYPE type, UInt32 flags);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool CredFree([In] IntPtr buffer);

        private enum CRED_TYPE : UInt32
        {
            GENERIC = 1,
            DOMAIN_PASSWORD = 2,
            DOMAIN_CERTIFICATE = 3,
            DOMAIN_VISIBLE_PASSWORD = 4,
            MAXIMUM = 5
        }

        private enum CRED_PERSIST : UInt32
        {
            SESSION = 1,
            LOCAL_MACHINE = 2,
            ENTERPRISE = 3
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private class CREDENTIAL
        {
            public UInt32 Flags;
            public CRED_TYPE Type;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string TargetName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public UInt32 CredentialBlobSize;
            public IntPtr CredentialBlob;
            public CRED_PERSIST Persist;
            public UInt32 AttributeCount;
            public IntPtr Attributes;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string TargetAlias;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string UserName;
        }

        #endregion PInvoke
    }
}
