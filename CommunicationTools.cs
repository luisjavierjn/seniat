/*******************************************************************************
$Id: //depot/Projects_Win_Rico/Seniat_net/Main/src/Seniat/Seniat/CommunicationTools.cs#2 $
$DateTime: 2011/07/28 11:36:56 $
$Change: 720 $
$Revision: #2 $
$Author: Zabel $
Copyright: ©2011 QUORiON Data Systems GmbH
*******************************************************************************/
using System;
using System.Text;

namespace QCom
{
    /// <summary>
    /// This static class is a universal toolbox for the byte and string handling
    /// </summary>
    public static class CommunicationTools
    {

        /// <summary>
        /// code page type
        /// </summary>
        public enum CodePageType
        {
            /// <summary>
            /// Western Europe
            /// </summary>
            WE, //Western Europe
            /// <summary>
            /// Central Europe
            /// </summary>
            CE, //Central Europe
            /// <summary>
            /// Baltic
            /// </summary>
            BA, //Baltic
            /// <summary>
            /// Grecian
            /// </summary>
            GR, //Grecian
            /// <summary>
            /// Arabian
            /// </summary>
            ARA, //Arabian
            /// <summary>
            /// Cyrillic
            /// </summary>
            CYS, //Cyrillic
            /// <summary>
            /// Hebrew
            /// </summary>
            HE, //Hebrew
        }


        /// <summary>
        /// convert a byte-array to a string
        /// </summary>
        /// <param name="data">byte array</param>
        /// <returns>string</returns>
        public static string byteArrayToString(byte[] data)
        {
            string sData = string.Empty;
            foreach (byte b in data)
            {
                sData = sData + Convert.ToChar(b);
            }
            return sData;
        }

        /// <summary>
        /// convert a byte to a string (ascii)
        /// </summary>
        /// <param name="sign">byte value</param>
        /// <returns>string as ascii-value</returns>
        public static string byteToString(byte sign)
        {
            return (Convert.ToChar(sign)).ToString();
        }

        /// <summary>
        /// convert a string to a byte array
        /// supported no unicode conversation
        /// </summary>
        /// <param name="data">string wich will be processed</param>
        /// <returns>byte array</returns>
        public static byte[] stringToByteArray(string data)
        {
            byte[] bArray = new byte[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                bArray[i] = Convert.ToByte(data[i]);
            }

            return bArray;
        }

        /// <summary>
        /// convert a byte array to a hex string
        /// </summary>
        /// <param name="data">byte array</param>
        /// <returns>a string which shows the bytes as hex values</returns>
        public static string ByteArrayToHexString(byte[] data)
        {
            StringBuilder sb = new StringBuilder(data.Length * 3);
            foreach (byte b in data)
                sb.Append(Convert.ToString(b, 16).PadLeft(2, '0').PadRight(3, ' '));
            return sb.ToString().ToUpper();
        }

        /// <summary>
        /// Convert a unicode String to a ASCII Byte Array for the selected codepage
        /// </summary>
        /// <param name="unicodeData">unicode string</param>
        /// <param name="codepage">convert to ASCII codepage</param>
        /// <returns>ASCII byte array for codepage</returns>
        public static byte[] unicodeStringToASCIIByteArray(string unicodeData, CodePageType codepage)
        {
            Encoding iso;

            //switch for another codepages
            switch (codepage)
            {
                case CodePageType.WE:
                    iso = Encoding.GetEncoding("windows-1252"); //Westeuropean 
                    break;
                case CodePageType.CE:
                    iso = Encoding.GetEncoding("windows-1250"); //Centraleuropean
                    break;
                case CodePageType.ARA:
                    iso = Encoding.GetEncoding("windows-1256");
                    break;
                case CodePageType.BA:
                    iso = Encoding.GetEncoding("windows-1257");
                    break;
                case CodePageType.CYS:
                    iso = Encoding.GetEncoding("windows-1251");
                    break;
                case CodePageType.GR:
                    iso = Encoding.GetEncoding("windows-1253");
                    break;
                case CodePageType.HE:
                    iso = Encoding.GetEncoding("windows-1255");
                    break;
                default:
                    iso = Encoding.GetEncoding("windows-1252");
                    break;
            }


            return iso.GetBytes(unicodeData);
        }
    }
}

