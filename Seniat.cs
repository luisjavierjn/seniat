/*******************************************************************************
$Id: //depot/Projects_Win_Rico/Seniat_net/Main/src/Seniat/Seniat/Seniat.cs#9 $
$DateTime: 2011/07/28 17:26:37 $
$Change: 722 $
$Revision: #9 $
$Author: Zabel $
Copyright: ©2011 QUORiON Data Systems GmbH
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using QCom;
using System.Threading;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;

namespace Seniat
{    
    /// <summary>
    /// Class for Seniat Communication
    /// </summary>
    public class Seniat
    {
        /// <summary>
        /// speed informations about the serial port
        /// </summary>
        public enum Baudrate
        {
            /// <summary>
            /// Baudrate 9600
            /// </summary>
            CBR_9600,
            /// <summary>
            /// Baudrate 19200
            /// </summary>
            CBR_19200,
            /// <summary>
            /// Baudrate 38400
            /// </summary>
            CBR_38400,
            /// <summary>
            /// Baudrate 57600
            /// </summary>
            CBR_57600,
            /// <summary>
            /// Baudrate 115200
            /// </summary>
            CBR_115200
        };
        
        
        private QCommunication qCom;
        private string _brandIdentificaton = string.Empty;
        private string _modelIdentification = string.Empty;
        private string _ret;
        private int _counterNAK = 0;
        private string _dateSince = string.Empty;
        private string _dateUntil = string.Empty;
        private string _lastError = string.Empty;
        QCommunication.BaudRate _baudrate = QCommunication.BaudRate.CBR_57600;

        /// <summary>
        /// set/get the communication baudrate
        /// </summary>
        public Baudrate BAUDRATE
        {
            get
            {
                switch (_baudrate)
                {
                    case QCommunication.BaudRate.CBR_9600:
                        return Baudrate.CBR_9600;

                    case QCommunication.BaudRate.CBR_19200:
                        return Baudrate.CBR_19200;

                    case QCommunication.BaudRate.CBR_57600:
                        return Baudrate.CBR_57600;

                    case QCommunication.BaudRate.CBR_38400:
                        return Baudrate.CBR_38400;

                    case QCommunication.BaudRate.CBR_115200:
                        return Baudrate.CBR_115200;

                    default:
                        _baudrate = QCommunication.BaudRate.CBR_57600;
                        return Baudrate.CBR_57600;

                }
            }

            set
            {
                switch (value)
                {
                    case Baudrate.CBR_9600:
                        _baudrate = QCommunication.BaudRate.CBR_9600;
                        break;

                    case Baudrate.CBR_19200:
                        _baudrate = QCommunication.BaudRate.CBR_19200;
                        break;

                    case Baudrate.CBR_38400:
                        _baudrate = QCommunication.BaudRate.CBR_38400;
                        break;

                    case Baudrate.CBR_57600:
                        _baudrate = QCommunication.BaudRate.CBR_57600;
                        break;
                        
                    case Baudrate.CBR_115200:
                        _baudrate = QCommunication.BaudRate.CBR_115200;
                        break;
                }
            }
        }

        /// <summary>
        /// get the last occur error message
        /// </summary>
        public string LastErrorMessage
        {
            get
            {
                return this._lastError;
            }
        }

        #region Public
        /// <summary>
        /// check connection
        /// </summary>
        /// <param name="brandIdentification">identification for the cr brand</param>
        /// <param name="modelIdentification">identification for the cr model</param>
        /// <param name="connectionPort">connection port</param>
        /// <returns>correct connection</returns>
        public bool verificarConexion(string brandIdentification, string modelIdentification, string connectionPort)
        {
            if (qCom != null)
            {
                qCom.Dispose();
                qCom = null;
            }

            if (qCom == null)
            {
                qCom = new QCommunication(connectionPort, _baudrate, 10000);
                qCom.OnCRCError += new QCommunication.QCommunicationCRCErrorHandler(qCom_OnCRCError);
                qCom.OnIncomingData += new QCommunication.QCommunicationReceiveHandler(qCom_OnIncomingData);
                qCom.OnTimeOut += new QCommunication.QCommunicationTimeOutHandler(qCom_OnTimeOut);
                qCom.OnACKMessage += new QCommunication.QCommunicationACKEvent(qCom_OnACKMessage);
                qCom.OnNAKMessage += new QCommunication.QCommunicationNAKEvent(qCom_OnNAKMessage);
                qCom.OnSYNMessage += new QCommunication.QCommunicationSYNEvent(qCom_OnSYNMessage);
            }

            _brandIdentificaton = brandIdentification;
            _modelIdentification = modelIdentification;
            

            ArrayList answerString = new ArrayList();
           
            _counterNAK = 0;
            string val = getInformation("F000000000000");

            
            while (!checkString(val) && _counterNAK < 3)
            {
                _counterNAK++;
                val = getNotAcknowlegeAnser();
            }

            if (_counterNAK >= 3)
            {
                setLastError(val);
                return false;
            }

            answerString.Add(val);

            while (!isRFRBlock(val))
            {
                    _counterNAK = 0;
                    val = getAcknowledgeAnswer();

                    while (!checkString(val) && _counterNAK < 5)
                    {
                        _counterNAK++;
                        val = getNotAcknowlegeAnser();
                        //qCom.sendNAK(false);
                    }

                    if (_counterNAK >= 5)
                    {
                        setLastError(val);
                        return false;
                    }

                    answerString.Add(val);
            }
            
            qCom.sendACK(false);
            qCom.Close();


#if DEBUG
            writeLog(answerString, "verificarConexion");
#endif

            int numberOfBlocks;

            if (!int.TryParse(answerString[answerString.Count - 1].ToString(), out numberOfBlocks))
            {
                setLastError("PNumBlocks");
                return false;
            }

            if (numberOfBlocks != answerString.Count - 1)
            {
                setLastError("NumBlocks");
                return false;
            }

            if (_brandIdentificaton != getBrand(answerString[0].ToString()))
            {
                setLastError("BrIdent");
                return false;
            }

            if (_modelIdentification != getModel(answerString[0].ToString()))
            {
                setLastError("modIdent");
                return false;
            }
            

            return true;
        }

        /// <summary>
        /// get header infos
        /// </summary>
        /// <param name="brandIdentification">identification for the cr brand</param>
        /// <param name="modelIdentification">identification for the cr model</param>
        /// <param name="exitFileName">saving filename</param>
        /// <returns>correct saving file</returns>
        public bool obtenerEncabezado(string brandIdentification, string modelIdentification, string exitFileName)
        {
            if (qCom != null)
            {
                if (File.Exists(exitFileName))
                    File.Delete(exitFileName);

                ArrayList answerString = new ArrayList();

                _counterNAK = 0;
                string val = getInformation("F000000000000");

                while (!checkString(val) && _counterNAK < 3)
                {
                    _counterNAK++;
                    val = getNotAcknowlegeAnser();
                }

                if (_counterNAK >= 3)
                {
                    setLastError(val);
                    return false;
                }

                answerString.Add(val);

                while (!isRFRBlock(val))
                {
                    _counterNAK = 0;
                    val = getAcknowledgeAnswer();

                    while (!checkString(val) && _counterNAK < 5)
                    {
                        _counterNAK++;
                        val = getNotAcknowlegeAnser();
                        //qCom.sendNAK(false);
                    }

                    if (_counterNAK >= 5)
                    {
                        setLastError(val);
                        return false;
                    }

                    answerString.Add(val);

                }

                qCom.sendACK(false);
                qCom.Close();


                int numberOfBlocks;

                if (!int.TryParse(answerString[answerString.Count - 1].ToString(), out numberOfBlocks))
                {
                    setLastError("PNumBlocks");
                    return false;
                }


                if (numberOfBlocks != answerString.Count - 1)
                {
                    setLastError("NumBlocks");
                    return false;
                }

                for (int i = 0; i < answerString.Count; i++)
                {
                    if (answerString[i].ToString().Length > 20)
                    {
                        if (_brandIdentificaton != getBrand(answerString[i].ToString()))
                        {
                            setLastError("BrIdent");
                            return false;
                        }
                        else if (_modelIdentification != getModel(answerString[i].ToString()))
                        {
                            setLastError("modIdent");
                            return false;
                        }
                        else
                            break;
                    }
                }

                string identString = string.Empty;

                if (answerString[0].ToString().Length > 20)
                    identString = answerString[0].ToString().Substring(0, 20);

                string fiscalIdentification = getFiscalIdentification(identString);
                string registerIdentification = getRegisterNummer(identString);

                if (fiscalIdentification == string.Empty)
                {
                    setLastError("fisIdent");
                    return false;
                }
                else if (registerIdentification == string.Empty)
                {
                    setLastError("regIdent");
                    return false;
                }

#if DEBUG
                writeLog(answerString, "obtenerEncabezado");
#endif

                StreamWriter sw = null;
                try
                {
                    sw = new StreamWriter(exitFileName, false);

                    sw.WriteLine("RIF Del Usuario\tNúmero Registro De La Maquina");
                    sw.WriteLine(fiscalIdentification + "\t" + registerIdentification);

                    sw.Close();
                }
                catch
                {
                    setLastError("WriteError");
                    return false;
                }

                return true;
            }
            else
            {
                setLastError("COMError");
                return false;
            }
        }

        /// <summary>
        /// get fiscal infos
        /// </summary>
        /// <param name="brandIdentification">identification for the cr brand</param>
        /// <param name="modelIdentification">identification for the cr model</param>
        /// <param name="DateSince">start date</param>
        /// <param name="DateUntil">stop date</param>
        /// <param name="exitFileName">saving filename</param>
        /// <returns></returns>
        public bool leerMemoriaFiscal(string brandIdentification, string modelIdentification, string DateSince, string DateUntil, string exitFileName)
        {
            if (qCom != null)
            {
                if (File.Exists(exitFileName))
                    File.Delete(exitFileName);

                ArrayList answerString = new ArrayList();

                _counterNAK = 0;

                DateTime startDate = DateTime.Parse(DateSince);
                DateTime stopDate = DateTime.Parse(DateUntil);

                string start = startDate.Year.ToString().Substring(2, 2) + startDate.Month.ToString("D2") + startDate.Day.ToString("D2");
                string stop = stopDate.Year.ToString().Substring(2, 2) + stopDate.Month.ToString("D2") + stopDate.Day.ToString("D2");


                string val = getInformation("F" + start + stop);

                while (!checkString(val) && _counterNAK < 3)
                {
                    _counterNAK++;
                    val = getNotAcknowlegeAnser();
                }

                if (_counterNAK >= 3)
                {
                    setLastError(val);
                    return false;
                }

                answerString.Add(val);

                while (!isRFRBlock(val))
                {
                    _counterNAK = 0;
                    val = getAcknowledgeAnswer();

                    while (!checkString(val) && _counterNAK < 5)
                    {
                        _counterNAK++;
                        val = getNotAcknowlegeAnser();
                        //qCom.sendNAK(false);
                    }

                    if (_counterNAK >= 5)
                    {
                        setLastError(val);
                        return false;
                    }

                    answerString.Add(val);

                }
                qCom.sendACK(false);
                qCom.Close();

                int numberOfBlocks;

                if (!int.TryParse(answerString[answerString.Count - 1].ToString(), out numberOfBlocks))
                {
                    setLastError("PNumBlocks");
                    return false;
                }


                if (numberOfBlocks != answerString.Count - 1)
                {
                    setLastError("NumBlocks");
                    return false;
                }

                for (int i = 0; i < answerString.Count; i++)
                {
                    if (answerString[i].ToString().Length > 20)
                    {
                        if (_brandIdentificaton != getBrand(answerString[i].ToString()))
                        {
                            setLastError("BrIdent");
                            return false;
                        }
                        else if (_modelIdentification != getModel(answerString[i].ToString()))
                        {
                            setLastError("modIdent");
                            return false;
                        }
                        else
                            break;
                    }
                }

#if DEBUG
                writeLog(answerString, "leerMemoriaFiscal");
#endif

                StreamWriter sw = null;
                try
                {
                    sw = new StreamWriter(exitFileName, false);

                    sw.WriteLine("Número Del Reporte Z\tFecha De Grabación (Z)\tNumero Última Factura\tFecha De La Ultima Factura\tHora De La Ultima Factura\tMonto Exento o Exonerado\tTotal Base Imponible General\tTotal IVA Alicuota General\tTotal Base Imponible Reducida\tTotal IVA Alicuota Reducida\tTotal Base Imponible Adicional\tTotal IVA Alicuota Adicional\tTotal Devoluciones Exentas Y/O Exoneradas\tTotal Devoluciones Base Imponible General\tTotal Devoluciones IVA Alicuota General\tTotal Devoluciones Base Imponible Reducida\tTotal Devoluciones IVA Alicuota Reducida\tTotal Devoluciones Base Imponible Adicional\tTotal Devoluciones IVA Alicuota Adicional");

                    for (int i = 0; i < answerString.Count - 1; i++)
                    {
                        string line = getFiscalLineString(answerString[i].ToString());
                        if (line != string.Empty)
                            sw.WriteLine(line);
                    }

                    sw.Close();
                }
                catch
                {
                    setLastError("WriteError");
                    return false;
                }

                return true;
            }
            else
            {
                setLastError("COMError");
                return false;
            }
        }

        /// <summary>
        /// get document infos
        /// </summary>
        /// <param name="brandIdentification">identification for the cr brand</param>
        /// <param name="modelIdentification">identification for the cr model</param>
        /// <param name="DateSince">start date</param>
        /// <param name="DateUntil">stop date</param>
        /// <param name="exitFileName">saving filename</param>
        /// <returns></returns>
        public bool obtenerDatosDocumento(string brandIdentification, string modelIdentification, string DateSince, string DateUntil, string exitFileName)
        {
            _dateSince = DateSince;
            _dateUntil = DateUntil;

            if (qCom != null)
            {
                if (File.Exists(exitFileName))
                    File.Delete(exitFileName);

                ArrayList answerString = new ArrayList();

                _counterNAK = 0;

                DateTime startDate = DateTime.Parse(DateSince);
                DateTime stopDate = DateTime.Parse(DateUntil);

                string start = startDate.Year.ToString().Substring(2, 2) + startDate.Month.ToString("D2") + startDate.Day.ToString("D2");
                string stop = stopDate.Year.ToString().Substring(2, 2) + stopDate.Month.ToString("D2") + stopDate.Day.ToString("D2");


                string val = getInformation("A" + start + stop);

                while (!checkString(val) && _counterNAK < 3)
                {
                    _counterNAK++;
                    val = getNotAcknowlegeAnser();
                }

                if (_counterNAK >= 3)
                {
                    setLastError(val);
                    return false;
                }

                answerString.Add(val);

                while (!isRFRBlock(val))
                {
                    _counterNAK = 0;
                    val = getAcknowledgeAnswer();

                    while (!checkString(val) && _counterNAK < 5)
                    {
                        _counterNAK++;
                        val = getNotAcknowlegeAnser();
                        //qCom.sendNAK(false);
                    }

                    if (_counterNAK >= 5)
                    {
                        setLastError(val);
                        return false;
                    }

                    answerString.Add(val);

                }

                qCom.sendACK(false);
                qCom.Close();


                /*
                int numberOfBlocks;

                if (!int.TryParse(answerString[answerString.Count - 1].ToString(), out numberOfBlocks))
                    return false;

                if (numberOfBlocks != answerString.Count - 1)
                    return false;*/

                /*
                if (_brandIdentificaton != getBrand(answerString[0].ToString()))
                    return false;

                if (_modelIdentification != getModel(answerString[0].ToString()))
                    return false;
                 */
#if DEBUG
                writeLog(answerString, "obtenerDatosDocumento");
#endif

                StreamWriter sw = null;
                try
                {

                    sw = new StreamWriter(exitFileName, false, Encoding.Default);
                    string[] lines = getDatosDocumento(answerString);

                    for (int i = 0; i < lines.Length; i++)
                    {
                        sw.WriteLine(lines[i]);
                    }

                    sw.Close();
                }
                catch
                {
                    setLastError("WriteError");
                    return false;
                }

                return true;
            }
            else
            {
                setLastError("COMError");
                return false;
            }
        }

        /// <summary>
        /// get totals infos
        /// </summary>
        /// <param name="brandIdentification">identification for the cr brand</param>
        /// <param name="modelIdentification">identification for the cr model</param>
        /// <param name="DateSince">start date</param>
        /// <param name="DateUntil">stop date</param>
        /// <param name="exitFileName">saving filename</param>
        /// <returns></returns>
        public bool obtenerTotales(string brandIdentification, string modelIdentification, string DateSince, string DateUntil, string exitFileName)
        {
            _dateSince = DateSince;
            _dateUntil = DateUntil;

            if (qCom != null)
            {
                if (File.Exists(exitFileName))
                    File.Delete(exitFileName);

                ArrayList answerString = new ArrayList();

                _counterNAK = 0;

                DateTime startDate = DateTime.Parse(DateSince);
                DateTime stopDate = DateTime.Parse(DateUntil);

                string start = startDate.Year.ToString().Substring(2, 2) + startDate.Month.ToString("D2") + startDate.Day.ToString("D2");
                string stop = stopDate.Year.ToString().Substring(2, 2) + stopDate.Month.ToString("D2") + stopDate.Day.ToString("D2");


                string val = getInformation("A" + start + stop);

                while (!checkString(val) && _counterNAK < 3)
                {
                    _counterNAK++;
                    val = getNotAcknowlegeAnser();
                }

                if (_counterNAK >= 3)
                {
                    setLastError(val);
                    return false;
                }

                answerString.Add(val);

                while (!isRFRBlock(val))
                {
                    _counterNAK = 0;
                    val = getAcknowledgeAnswer();

                    while (!checkString(val) && _counterNAK < 5)
                    {
                        _counterNAK++;
                        val = getNotAcknowlegeAnser();
                        //qCom.sendNAK(false);
                    }

                    if (_counterNAK >= 5)
                    {
                        setLastError(val);
                        return false;
                    }

                    answerString.Add(val);

                }
                
                qCom.sendACK(false);
                qCom.Close();
 
                /*
                int numberOfBlocks;

                if (!int.TryParse(answerString[answerString.Count - 1].ToString(), out numberOfBlocks))
                    return false;

                if (numberOfBlocks != answerString.Count - 1)
                    return false;*/

                /*
                if (_brandIdentificaton != getBrand(answerString[0].ToString()))
                    return false;

                if (_modelIdentification != getModel(answerString[0].ToString()))
                    return false;
                 */

#if DEBUG
                writeLog(answerString, "obtenerTotales");
#endif

                StreamWriter sw = null;
                try
                {

                    sw = new StreamWriter(exitFileName, false, Encoding.Default);
                    string[] lines = this.getTotales(answerString);

                    for (int i = 0; i < lines.Length; i++)
                    {
                        sw.WriteLine(lines[i]);
                    }

                    sw.Close();
                }
                catch
                {
                    setLastError("WriteError");
                    return false;
                }



                return true;
            }
            else
            {
                setLastError("COMError");
                return false;
            }
        }

        /// <summary>
        /// get total infos 
        /// </summary>
        /// <param name="DocumentType">type of the document</param>
        /// <param name="DocumentNumber">number of the document</param>
        /// <param name="brandIdentification">identification for the cr brand</param>
        /// <param name="modelIdentification">identification for the cr model</param>
        /// <param name="exitFileName">saving filename</param>
        /// <returns></returns>
        public bool obtenerDetalleArticulos(string DocumentType, string DocumentNumber, string brandIdentification, string modelIdentification, string exitFileName)
        {
            if (qCom != null)
            {
                if (File.Exists(exitFileName))
                    File.Delete(exitFileName);

                ArrayList answerString = new ArrayList();

                _counterNAK = 0;

                if (_dateSince == string.Empty || _dateUntil == string.Empty)
                {
                    setLastError("DateError");
                    return false;
                }

                DateTime startDate = DateTime.Parse(_dateSince);
                DateTime stopDate = DateTime.Parse(_dateUntil);

                string start = startDate.Year.ToString().Substring(2, 2) + startDate.Month.ToString("D2") + startDate.Day.ToString("D2");
                string stop = stopDate.Year.ToString().Substring(2, 2) + stopDate.Month.ToString("D2") + stopDate.Day.ToString("D2");

                string val = getInformation("A" + start + stop);

                while (!checkString(val) && _counterNAK < 3)
                {
                    _counterNAK++;
                    val = getNotAcknowlegeAnser();
                }

                if (_counterNAK >= 3)
                {
                    setLastError(val);
                    return false;
                }

                answerString.Add(val);

                while (!isRFRBlock(val))
                {
                    _counterNAK = 0;
                    val = getAcknowledgeAnswer();

                    while (!checkString(val) && _counterNAK < 5)
                    {
                        _counterNAK++;
                        val = getNotAcknowlegeAnser();
                        //qCom.sendNAK(false);
                    }

                    if (_counterNAK >= 5)
                    {
                        setLastError(val);
                        return false;
                    }

                    answerString.Add(val);

                }

                qCom.sendACK(false);
                qCom.Close();


                /*
                int numberOfBlocks;

                if (!int.TryParse(answerString[answerString.Count - 1].ToString(), out numberOfBlocks))
                    return false;

                if (numberOfBlocks != answerString.Count - 1)
                    return false;*/

                /*
                if (_brandIdentificaton != getBrand(answerString[0].ToString()))
                    return false;

                if (_modelIdentification != getModel(answerString[0].ToString()))
                    return false;
                 */
#if DEBUG
                writeLog(answerString, "obtenerDetalleArticulos");
#endif

                StreamWriter sw = null;
                try
                {
                    sw = new StreamWriter(exitFileName, false, Encoding.Default);
                    string[] lines = this.getArticle(answerString, DocumentType, DocumentNumber);

                    if (lines.Length > 0)
                    {
                        for (int i = 0; i < lines.Length; i++)
                        {
                            sw.WriteLine(lines[i]);
                        }
                    }

                    sw.Close();
                }
                catch
                {
                    setLastError("WriteError");
                    return false;
                }
                return true;
            }
            else
            {
                setLastError("COMError");
                return false;
            }
        } 
        #endregion

        #region Private
        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        private string getInformation(string command)
        {
            if (qCom != null)
            {
                _ret = string.Empty;
                qCom.SendCommand(command, true);


                while (_ret == string.Empty || _ret == "SYN")
                {
                    if (_ret == "SYN")
                    {
                        qCom.resetSerialTimeoutValue();
                        _ret = string.Empty;
                    }
                    
                    Thread.Sleep(10);
                }
                if (_ret == "NAK" && _counterNAK < 3)
                {
                    
                    _ret = getInformation(command);
                    _counterNAK++;
                }

                return _ret;
            }
            else
                return "ERROR";
        }

        private string getAcknowledgeAnswer()
        {
            if (qCom != null)
            {
                _ret = string.Empty;
                qCom.sendACK(true);

                while (_ret == string.Empty || _ret == "SYN")
                {
                    if (_ret == "SYN")
                    {
                        qCom.resetSerialTimeoutValue();
                        _ret = string.Empty;
                    }

                    Thread.Sleep(10);
                }
                if (_ret == "NAK" && _counterNAK < 3)
                {

                    _ret = getAcknowledgeAnswer();
                    _counterNAK++;
                }

                return _ret;
            }
            else
                return "ERROR";
        }

        private string getNotAcknowlegeAnser()
        {
            if (qCom != null)
            {
                _ret = string.Empty;
                qCom.sendNAK(true);

                while (_ret == string.Empty || _ret == "SYN")
                {
                    if (_ret == "SYN")
                    {
                        qCom.resetSerialTimeoutValue();
                        _ret = string.Empty;
                    }

                    Thread.Sleep(10);
                }
                if (_ret == "NAK" && _counterNAK < 3)
                {

                    _ret = getAcknowledgeAnswer();
                    _counterNAK++;
                }

                return _ret;
            }
            else
                return "ERROR";
        }

        private bool checkString(string value)
        {
            if (value == "ERROR" || value == "CRC-Error" || value == "TIMEOUT")
                return false;
            else
                return true;
        }

        private void setLastError(string value)
        {
            _lastError = string.Empty;
            switch (value)
            {
                case "ERROR":
                    _lastError = "Unknown Error";
                    break;
                case "CRC-Error":
                    _lastError = "CRC Error";
                    break;
                case "TIMEOUT":
                    _lastError = "Timeout";
                    break;
                case "WriteError":
                    _lastError = "Can´t write File!";
                    break;
                case "COMError":
                    _lastError = "Invalid Communication!";
                    break;
                case "PNumBlocks":
                    _lastError = "Parsing Number of Blocks Error!";
                    break;
                case "NumBlocks":
                    _lastError = "Invalid Number of Blocks!";
                    break;
                case "BrIdent":
                    _lastError = "Invalid Brand Identification!";
                    break;
                case "modIdent":
                    _lastError = "Invalid Model Identification!";
                    break;
                case "fisIdent":
                    _lastError = "Fiscal Identification is empty";
                    break;
                case "regIdent":
                    _lastError = "Register Identification is empty";
                    break;
                case "DateError":
                    _lastError = "Invalid Start or End Datetime";
                    break;
                default:
                    _lastError = "Invalid Error";
                    break;
            }
        }

        private bool isRFRBlock(string value)
        {
            if (value.Length < 10)
            {
                int iValue;
                if (int.TryParse(value, out iValue))
                    return true;
            }

            return false;
        }

        private string getModel(string value)
        {
            if (value.Length > 15)
            {
                return value.Substring(12, 1);
            }
            else
                return string.Empty;
        }

        private string getBrand(string value)
        {
            if (value.Length > 15)
            {
                return value.Substring(11, 1);
            }
            else
                return string.Empty;
        }

        private string getFiscalIdentification(string value)
        {
            if (value.Length > 10)
            {
                return value.Substring(0, 10);
            }
            else
                return string.Empty;
        }

        private string getRegisterNummer(string value)
        {
            if (value.Length < 21)
            {
                return value.Substring(10, 10);
            }
                return string.Empty;
        }

        private string getFiscalLineString(string value)
        {
            if (value.Length > 20)
            {
                string ret = string.Empty;
                value = value.Substring(20);
                int length = value.Length;
                int[] width = new int[]{    4,
                                        10,
                                        8,
                                        10,
                                        5,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15,
                                        15};
                int pos = 0;

                for (int i = 0; i < width.Length; i++)
                {
                    if ((pos + width[i]) > value.Length)
                    {
                        break;
                    }
                    if (i > 4)
                    {
                        double dVal;
                        if (double.TryParse(value.Substring(pos, width[i]), out dVal))
                        {
                            ret += dVal.ToString("F2") + "\t";
                        }
                    }
                    else
                        ret += value.Substring(pos, width[i]) + "\t";
                    pos += width[i];
                }
                return ret;
                //throw new NotImplementedException();
            }
            else
                return string.Empty;
        }

        private string[] getDatosDocumento(ArrayList list)
        {
            ArrayList infoList = new ArrayList();
            infoList.Add("Tipo de Documento\tNúmero de Documento\tFecha de emisión\tHora de Emisión\t");
            string tmpline = string.Empty;
            string info = string.Empty;
            string writeLine = string.Empty;

            foreach (string ln in list)
            {
                string line = ln;

                if(tmpline.Length > 0)
                    line = tmpline.Substring(tmpline.Length - 3) + ln;
               
                tmpline = ln;

                for (int i = 0; i < line.Length; i++)
                {
                    if (line[i] != '\r' && line[i] != '\n')
                    {
                        info += QCom.CommunicationTools.byteToString((byte)line[i]);
                    }
                    else if (line[i] == '\n' && info.Length > 0)
                    {
                        byte controlSign = (byte)line[i + 2];

                        if (controlSign == 0xF4) //Dokumentennummer
                        {
                            if (info[0] == '0')
                                info = info.Substring(2);
                            string[] sub = info.Trim().Split(' ');
                            for (int j = 0; j < sub.Length; j++)
                            {
                                if (sub[j].Length > 0)
                                {
                                    int num;
                                    if (int.TryParse(sub[j], out num))
                                    {
                                        writeLine = " \t" + sub[j];
                                        break;
                                    }
                                }
                            }
                        }

                        if (controlSign == 0xF3)  // Datum/Uhrzeit
                        {
                            string[] sub = info.Substring(2).Trim().Split(' ');
                            for (int j = 0; j < sub.Length; j++)
                            {
                                if (sub[j].Length > 0)
                                {
                                    writeLine += "\t" + sub[j];
                                    //break;
                                }
                            }
                        }

                        if (controlSign == 0xF1) //Typ Dokument
                        {
                            string[] sub = info.Substring(2).Trim().Split(' ');
                            for (int j = 0; j < sub.Length; j++)
                            {
                                if (sub[j].Length > 0)
                                {
                                    writeLine = sub[j] + writeLine ;
                                }
                            }
                            if (writeLine.Length > 0)
                            {
                                infoList.Add(writeLine);
                            }
                            writeLine = string.Empty;
                        }

                        info = string.Empty;
                    }
                }
            }

            string[] lines = new string[infoList.Count];
            int a = 0;

            foreach (string ln in infoList)
            {
                lines[a] = ln;
                a++;
            }

            return lines;
        }

        private string[] getTotales(ArrayList list)
        {
            ArrayList infoList = new ArrayList();
            infoList.Add("Total Base Imponible General\tTotal Iva AlícuotaGeneral\tTotal Base Imponible Reducida\tTotal Iva AlícuotaReducida\tTotal Base Imponible Adicional\tTotal Iva AlícuotaAdicional");
            string tmpline = string.Empty;
            string info = string.Empty;

            int[] tmpValues = null;
            int[] totalValues = new int[6];
            int[] totalCreditoValues = new int[6];
            int taxCounter = 0;
            bool isCreditoValue = false;

            foreach (string ln in list)
            {
                string line = ln;

                if (tmpline.Length > 0)
                    line = tmpline.Substring(tmpline.Length - 3) + ln;

                tmpline = ln;

                for (int i = 0; i < line.Length; i++)
                {
                    if (line[i] != '\r' && line[i] != '\n')
                    {
                        info += QCom.CommunicationTools.byteToString((byte)line[i]);
                    }
                    else if (line[i] == '\n' && info.Length > 0)
                    {
                        byte controlSign = (byte)line[i + 2];
                        

                        if (controlSign == 0xF4) //new Receipt
                        {
                            if (tmpValues != null)
                            {
                                if (!isCreditoValue)  //normal value
                                {
                                    for (int j = 0; j < tmpValues.Length; j++)
                                        totalValues[j] += tmpValues[j];
                                }
                                else  //credit value
                                {
                                    for (int j = 0; j < tmpValues.Length; j++)
                                        totalCreditoValues[j] += tmpValues[j];
                                }
                            }

                            tmpValues = new int[6];
                            taxCounter = 0;
                            isCreditoValue = false;
                        }

                        if (controlSign == 0xF1) // Typ Dokument
                        {
                            string[] sub = info.Substring(2).Trim().Split(' ');
                            for (int j = 0; j < sub.Length; j++)
                            {
                                if (sub[j].Length > 0)
                                {
                                    int docType;
                                    if (int.TryParse(sub[j], out docType))
                                    {
                                        if (docType == 3)
                                            isCreditoValue = true;  //filter credit payment for extra counter - now will this will value not write in file
                                    }
                                }
                            }
                        }

                        if (controlSign == 0xF2) //tax values
                        {
                            string[] sub = info.Substring(2).Trim().Split(' ');
                            int subCounter = 0;

                            for (int j = 0; j < sub.Length; j++)
                            {
                                if(subCounter < 2)
                                {
                                    if (sub[j].Length > 0)
                                    {
                                        subCounter++;
                                        int val;
                                        if (int.TryParse(sub[j],out  val))
                                        {
                                            switch (taxCounter)
                                            {
                                                case 0:
                                                    tmpValues[taxCounter] = val;
                                                    break;

                                                case 1:
                                                    tmpValues[taxCounter] = val;
                                                    break;

                                                case 2:
                                                    tmpValues[taxCounter] = val;
                                                    break;

                                                case 3:
                                                    tmpValues[taxCounter] = val;
                                                    break;


                                                case 4:
                                                    tmpValues[taxCounter] = val;
                                                    break;

                                                case 5:
                                                    tmpValues[taxCounter] = val;
                                                    break;
                                            }
                                            taxCounter++;
                                        }
                                    }
                                }
                            }
                        }

                        info = string.Empty;
                    }
                }
            }

            if (tmpValues != null && taxCounter != 0)
            {
                if(!isCreditoValue)  //normal value
                {
                    for (int j = 0; j < tmpValues.Length; j++)
                        totalValues[j] += tmpValues[j];
                }
                else   //credit value
                {
                    for (int j = 0; j < tmpValues.Length; j++)
                        totalCreditoValues[j] += tmpValues[j];
                }
            }

            infoList.Add(this.getValuesLine(totalValues));
            //infoList.Add(this.getValuesLine(totalCreditoValues));  //not activated yet....for further using


            string[] lines = new string[infoList.Count];
            int a = 0;

            foreach (string ln in infoList)
            {
                lines[a] = ln;
                a++;
            }

            return lines;
        }

        string getValuesLine(int[] values)
        {
            string valline = string.Empty;
            for (int i = 0; i < values.Length; i++)
            {
                if (i > 2 && values[i] <= 0)
                {
                    valline += "\t";
                    break;
                }

                string sVal = values[i].ToString();
                int length = sVal.Length;
                if (length < 3)
                {
                    switch (length)
                    {
                        case 0:
                            sVal = "0,00";
                            break;

                        case 1:
                            sVal = "0,0" + sVal;
                            break;

                        case 2:
                            sVal = "0," + sVal;
                            break;
                    }

                }
                else
                    sVal = sVal.Insert(sVal.Length - 2, ",");

                valline += sVal + "\t";
            }
            return valline;
        }

        private string[] getArticle(ArrayList list, string documentType, string documentNumber)
        {
            ArrayList infoList = new ArrayList();
            infoList.Add("Nombre del Articulo\tPrecio");
            string tmpline = string.Empty;
            string info = string.Empty;
            string article = string.Empty;

            int docNum;
            bool docFound = false;
            bool docType = false;

            if (!int.TryParse(documentNumber, out docNum))
                return null;

            foreach (string ln in list)
            {
                string line = ln;

                if (tmpline.Length > 0)
                    line = tmpline.Substring(tmpline.Length - 3) + ln;

                tmpline = ln;

                for (int i = 0; i < line.Length; i++)
                {
                    if (line[i] != '\r' && line[i] != '\n')
                    {
                        info += QCom.CommunicationTools.byteToString((byte)line[i]);
                    }
                    else if (line[i] == '\n' && info.Length > 0)
                    {
                        byte controlSign = (byte)line[i + 2];

                        if (controlSign == 0xF4) //new Receipt
                        {
                            if (info[0] == '0')
                                info = info.Substring(2);
                            string[] sub = info.Trim().Split(' ');
                            for (int j = 0; j < sub.Length; j++)
                            {
                                if (sub[j].Length > 0)
                                {
                                    int num;
                                    if (int.TryParse(sub[j], out num))
                                    {
                                        if (num == docNum)
                                            docFound = true;
                                        else
                                            docFound = false;
                                    }
                                }
                            }
                        }

                        if (controlSign == 0xF1) //documenttype
                        {
                            string[] sub = info.Substring(2).Trim().Split(' ');
                            for (int j = 0; j < sub.Length; j++)
                            {
                                if (sub[j].Length > 0)
                                {
                                    if (sub[j] == documentType)
                                        docType = true;
                                    else
                                        docType = false;
                                }
                            }
                        }

                        if (docFound && docType)
                        {
                            if (controlSign == 0xF6 || controlSign == 0xF8) //article name || price
                            {
                                string sub = info.Substring(2).Trim();
                                if (sub.Length > 0)
                                {

                                    if (controlSign == 0xF8) //name
                                    {
                                        article = sub;
                                    }

                                    if (controlSign == 0xF6) //price
                                    {
                                        if (sub.Trim().EndsWith("%") || sub.Trim().EndsWith("(E)"))
                                        {
                                            article = sub.Substring(0, sub.IndexOf("Bs."));
                                            string tmpPrice = sub.Substring(sub.IndexOf("Bs."));
                                            char[] separators = { ' ' };
                                            string[] parts = tmpPrice.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                                            if (parts.Length == 2)
                                            {
                                                article += parts[1];
                                            }
                                            if (parts.Length > 0)
                                            {
                                                string[] prices = ExtractNumbers(parts[0], false);
                                                if (prices.Length > 0)
                                                {
                                                    if (prices.Length == 2)
                                                    {
                                                        article += parts[0] + "\t" + prices[1];
                                                    }
                                                    else
                                                        article += "\t" + prices[0];
                                                }
                                            }

                                        }
                                        else
                                        {
                                            char[] separators = { ' ' };
                                            string[] parts = sub.Split(separators, StringSplitOptions.RemoveEmptyEntries);

                                            string[] prices = ExtractNumbers(sub, false);
                                            if (prices.Length > 0)
                                            {
                                                if (prices.Length == 2)
                                                {
                                                    article += parts[0] + "\t" + prices[1];
                                                }
                                                else
                                                    article += "\t" + prices[0];
                                            }
                                        }
                                        infoList.Add(article);
                                        article = string.Empty;
                                    }
                                }
                            }
                        }

                        info = string.Empty;
                    }
                }
            }


            string[] lines = new string[infoList.Count];
            int a = 0;

            foreach (string ln in infoList)
            {
                lines[a] = ln;
                a++;
            }

            return lines;
        }

        /// <summary>
        /// extract numbers from a source string with pattern
        /// </summary>
        /// <param name="source">source string with including numbers</param>
        /// <param name="extractOnlyIntegers">extract numbers only as integer - decimal separators will be ignore</param>
        /// <returns>list of extract numbers as string</returns>
        private string[] ExtractNumbers(string source, bool extractOnlyIntegers)
        {
            string pattern;
            if (extractOnlyIntegers)
            {
                pattern = @"-?\d{1,}";
            }
            else
            {
                string decimalSeparator = ",";
                pattern = @"-?\d{1,}" + decimalSeparator + @"{0,1}\d{0,}";
            }

            MatchCollection matches = Regex.Matches(source, pattern);

            string[] result = new string[matches.Count];
            for (int i = 0; i < matches.Count; i++)
                result[i] = matches[i].Value;

            return result;
        }

#if DEBUG
        private void writeLog(ArrayList stringList, string title)
        {
            string filename = "inLog.log";
            StreamWriter sw = new StreamWriter(filename, true);
            DateTime now = DateTime.Now;
            sw.WriteLine();
            sw.WriteLine(now.ToLongDateString() + " " + now.ToLongTimeString() + "," + now.Millisecond.ToString("000"));
            sw.WriteLine(title);
            sw.WriteLine("---BEGIN---");

            foreach (string s in stringList)
            {
                sw.Write(s);
            }

            sw.WriteLine();
            sw.WriteLine("---END---");

            sw.Close();
        }

        private void writeErrorLog(string error, string message, string title)
        {
            try
            {
                string filename = "errorLog.log";
                StreamWriter sw = new StreamWriter(filename, true);
                DateTime now = DateTime.Now;
                sw.WriteLine(now.ToLongDateString() + " " + now.ToLongTimeString() + "," + now.Millisecond.ToString("000"));
                sw.WriteLine(title + ": " + error);
                sw.WriteLine(message);
                sw.WriteLine();

                sw.Close();
            }
            catch{}
        }
#endif
        #endregion

#region Events
        void qCom_OnTimeOut(object sender, EventArgs e)
        {
            _ret = "TIMEOUT";
#if DEBUG
            writeErrorLog("EventMessage: " + _ret, "Message for Timeout not available", "ERROR-EVENT");
#endif
            //throw new NotImplementedException();
        }

        void qCom_OnIncomingData(object sender, QCommunicationEventArgs e)
        {
            _ret = e.ReceiveMessageString;
            //throw new NotImplementedException();
        }

        void qCom_OnCRCError(object sender, QCommunicationCRCEventArgs e)
        {
            _ret = "CRC-Error";
#if DEBUG
            StringBuilder message = new StringBuilder();
            message.Append("Message: ");
            message.Append(CommunicationTools.ByteArrayToHexString(e.ReceiveMessageByte));
            message.Append("\n");
            message.Append("Lenght: " + e.ReceiveMessageByte.Length.ToString() + "\n");
            message.Append("Checksum of the Message: ");
            message.Append(CommunicationTools.ByteArrayToHexString(new byte[] {e.IncomingCRCValue}));
            message.Append("\n");
            message.Append("Excepted Checksum: ");
            message.Append(CommunicationTools.ByteArrayToHexString(new byte[] {e.ExceptCRCValue}));
            message.Append("\n");
            writeErrorLog("EventMessage: " + _ret,message.ToString(), "ERROR-EVENT");
#endif
            //throw new NotImplementedException();
        }

        void qCom_OnSYNMessage(object sender, EventArgs e)
        {
            _ret = "SYN";
            //throw new NotImplementedException();
        }

        void qCom_OnNAKMessage(object sender, EventArgs e)
        {
            _ret = "NAK";
#if DEBUG
            writeErrorLog("EventMessage: " + _ret, "Message for NAK not available!", "ERROR-EVENT");
#endif
            //throw new NotImplementedException();
        }

        void qCom_OnACKMessage(object sender, EventArgs e)
        {
            _ret = "ACK";
            //throw new NotImplementedException();
        }
        #endregion

    }
}
