using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ExcelDna.Integration;
using System.Net.Sockets;
using System.Net;
using System.Net.NetworkInformation;
using DnsClient;
using DnsClient.Protocol;

namespace IPTools
{
    public partial class ThisAddIn
    {
        public static string GetDNS([ExcelArgument(Name = "GetDNSAddress", Description = "ReverseDNS from IP Address ")] string GetDNSAddress)
        {
            try
            {
                IPHostEntry hostInfo = Dns.Resolve(GetDNSAddress);

                return hostInfo.HostName;
                
            }
            catch (SocketException e)
            {
                return "SocketException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (ArgumentNullException e)
            {
                return "ArgumentNullException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (NullReferenceException e)
            {
                return "NullReferenceException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (Exception e)
            {
                return "Exception caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }


        }

        public static string GetDNSWSERVER([ExcelArgument(Name = "GetDNSAddress", Description = "ReverseDNS from IP Address ")] string GetDNSAddress, [ExcelArgument(Name = "DNSServer", Description = "DNS Server Address ")] string DNSServer)
        {
            try
            {
                var endpoint = new IPEndPoint(IPAddress.Parse(DNSServer), 53);
                var lookup = new LookupClient(endpoint);


                return lookup.QueryReverse(IPAddress.Parse(GetDNSAddress)).Answers.PtrRecords().FirstOrDefault().PtrDomainName;
            }
            catch (SocketException e)
            {
                return "SocketException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (ArgumentNullException e)
            {
                return "ArgumentNullException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (NullReferenceException e)
            {
                return "NullReferenceException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (Exception e)
            {
                return "Exception caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }


        }

        public static string GetIP([ExcelArgument(Name = "GetIPAddress", Description = "IP Adress of Host")] string GetIPAddress)
        {
            try
            {

                    IPHostEntry hostInfo = Dns.GetHostByName(GetIPAddress);
                    return hostInfo.AddressList[0].ToString();


            }
            catch (SocketException e)
            {
                return "SocketException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (ArgumentNullException e)
            {
                return "ArgumentNullException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (NullReferenceException e)
            {
                return "NullReferenceException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (Exception e)
            {
                return "Exception caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }


        }

        public static string GetIPWSERVER([ExcelArgument(Name = "GetIPAddress", Description = "IP Adress of Host")] string GetIPAddress, [ExcelArgument(Name = "DNSServer", Description = "DNS Server Address ")] string DNSServer)
        {
            try
            {

                var endpoint = new IPEndPoint(IPAddress.Parse(DNSServer), 53);
                var lookup = new LookupClient(endpoint);

                //return lookup.Query(GetIPAddress, QueryType.A).Answers.ARecords().FirstOrDefault().Address.ToString();
                return lookup.GetHostEntry(GetIPAddress).AddressList[0].ToString();



            }
            catch (SocketException e)
            {
                return "SocketException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (ArgumentNullException e)
            {
                return "ArgumentNullException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (NullReferenceException e)
            {
                return "NullReferenceException caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }
            catch (Exception e)
            {
                return "Exception caught!!! , Source : " + e.Source + "Message : " + e.Message;
            }


        }

        public static bool PING(string PINGAddress)
        {
            bool pingable = false;
            Ping pinger = null;

            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(PINGAddress);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (SocketException e)
            {
                return pingable;
            }
            catch (ArgumentNullException e)
            {
                return pingable;
            }
            catch (NullReferenceException e)
            {
                return pingable;
            }
            catch (Exception e)
            {
                return pingable;
            }
            return pingable;

        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
        
        
        
        

}
}
