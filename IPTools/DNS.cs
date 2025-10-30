
/*

Part of this code was written by @Jaimon Mathew. Thanks Jaimon for your contributoin!

Changes and improvements were made by <seixaserick77@gmail.com> to optimize the code for .NET 6.0 and implement some coding conventions.

 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace DnsLib
{

    public class DnsLite
    {
        private byte[] data;
        private int position, id, length;
        private string name;
        private string dnsServers;

        private static int DNS_PORT = 53; //DNS protocol uses UDP 53 port by Default

        Encoding ASCII = Encoding.ASCII;

        public DnsLite()
        {
            id = DateTime.Now.Millisecond * 60;
        }

        public void setDnsServers(string dnsServers)
        {
            this.dnsServers = dnsServers;
        }
        
        private int getNewId()
        {
            //return a new id
            return ++id;
        }

        //for packing the information to the format accepted by server
        public void MakeQuery(int id, String name)
        {

            data = new byte[512];

            for (int i = 0; i < 512; ++i)
            {
                data[i] = 0;
            }

            data[0] = (byte)(id >> 8);
            data[1] = (byte)(id & 0xFF);
            data[2] = (byte)1; data[3] = (byte)0;
            data[4] = (byte)0; data[5] = (byte)1;
            data[6] = (byte)0; data[7] = (byte)0;
            data[8] = (byte)0; data[9] = (byte)0;
            data[10] = (byte)0; data[11] = (byte)0;

            string[] tokens = name.Split(new char[] { '.' });
            string label;

            position = 12;

            for (int j = 0; j < tokens.Length; j++)
            {

                label = tokens[j];
                data[position++] = (byte)(label.Length & 0xFF);
                byte[] b = ASCII.GetBytes(label);

                for (int k = 0; k < b.Length; k++)
                {
                    data[position++] = b[k];
                }

            }

            data[position++] = (byte)0; data[position++] = (byte)0;
            data[position++] = (byte)15; data[position++] = (byte)0;
            data[position++] = (byte)1;

        }

        
        private int proc(int position)
        {

            int len = (data[position++] & 0xFF);

            if (len == 0)
            {
                return position;
            }

            int offset;

            do
            {
                if ((len & 0xC0) == 0xC0)
                {
                    if (position >= length)
                    {
                        return -1;
                    }
                    offset = ((len & 0x3F) << 8) | (data[position++] & 0xFF);
                    proc(offset);
                    return position;
                }
                else
                {
                    if ((position + len) > length)
                    {
                        return -1;
                    }
                    name += ASCII.GetString(data, position, len);
                    position += len;
                }

                if (position > length)
                {
                    return -1;
                }

                len = data[position++] & 0xFF;

                if (len != 0)
                {
                    name += ".";
                }
            } while (len != 0);

            return position;
        }
    }
}
