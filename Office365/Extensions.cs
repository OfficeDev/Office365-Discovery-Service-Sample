﻿//----------------------------------------------------------------------------------------------
//    Copyright 2014 Microsoft Corporation
//
//    Licensed under the Apache License, Version 2.0 (the "License");
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
//
//      http://www.apache.org/licenses/LICENSE-2.0
//
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
//----------------------------------------------------------------------------------------------

using Windows.Storage.Streams;

namespace Win8ServiceDiscovery
{
    static class Extensions
    {
        public static void WriteStringWithLength(this DataWriter w, string s)
        {
            w.WriteUInt32((uint)s.Length);
            w.WriteString(s);
        }

        public static string ReadString(this DataReader r)
        {
            return r.ReadString(r.ReadUInt32());
        }
    }
}
