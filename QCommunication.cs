/*******************************************************************************
$Id: //depot/Projects_Win_Rico/Seniat_net/Main/src/Seniat/Seniat/QCommunication.cs#4 $
$DateTime: 2011/07/28 14:07:18 $
$Change: 721 $
$Revision: #4 $
$Author: Zabel $
Copyright: ©2011 QUORiON Data Systems GmbH
*******************************************************************************/
using System;
using System.Collections;

namespace QCom
{
    /// <summary>
    /// Class for Communication to the QUORiON Devices
    /// </summary>
    public class QCommunication : IDisposable
    {
        #region Structs/Enum
        /// <summary>
        /// Enumerate the possible Baudrates
        /// </summary>
        public enum BaudRate : int
        {
            /// <summary>
            /// Baudrate 9600 bits per second
            /// </summary>
            CBR_9600,
            /// <summary>
            /// Baudrate 19200 bits per second
            /// </summary>
            CBR_19200,
            /// <summary>
            /// Baudrate 38400 bits per second
            /// </summary>
            CBR_38400,
            /// <summary>
            /// Baudrate 57600 bits per second
            /// </summary>
            CBR_57600,
            /// <summary>
            /// Baudrate 115200 bits per second
            /// </summary>
            CBR_115200
        }
        #endregion

        #region EventHandler
        /// <summary>
        /// Handle for incoming Data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public delegate void QCommunicationReceiveHandler(object sender, QCommunicationEventArgs e);
        /// <summary>
        /// Handle for timeout
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public delegate void QCommunicationTimeOutHandler(object sender, EventArgs e);
        /// <summary>
        /// Handle for CRC Error
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public delegate void QCommunicationCRCErrorHandler(object sender, QCommunicationCRCEventArgs e);
       /// <summary>
       /// Handle for NAK Signal
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        public delegate void QCommunicationNAKEvent(object sender, EventArgs e);
        /// <summary>
        /// Handle for ACK Signal
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public delegate void QCommunicationACKEvent(object sender, EventArgs e);
        /// <summary>
        /// Handle for SYN Signal
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public delegate void QCommunicationSYNEvent(object sender, EventArgs e);

        /// <summary>
        /// Event for incoming data signal
        /// </summary>
        public event QCommunicationReceiveHandler OnIncomingData;
        /// <summary>
        /// Event for timeout signal
        /// </summary>
        public event QCommunicationTimeOutHandler OnTimeOut;
        /// <summary>
        /// Event for a CRC Error signal
        /// </summary>
        public event QCommunicationCRCErrorHandler OnCRCError;
        /// <summary>
        /// Event for a ACK Message
        /// </summary>
        public event QCommunicationACKEvent OnACKMessage;
        /// <summary>
        /// Event for a NAK Message
        /// </summary>
        public event QCommunicationNAKEvent OnNAKMessage;
        /// <summary>
        /// Event for a SYN Message
        /// </summary>
        public event QCommunicationSYNEvent OnSYNMessage;
        #endregion

        #region Control Charakter
        /// <summary>
        /// Null
        /// </summary>
        const byte _cc_NUL = 0x00;
        /// <summary>
        /// start of heading
        /// </summary>
        const byte _cc_SOH = 0x01;
        /// <summary>
        /// start of text
        /// </summary>
        const byte _cc_STX = 0x02;
        /// <summary>
        /// end of text
        /// </summary>
        const byte _cc_ETX = 0x03;
        /// <summary>
        /// end of transmission
        /// </summary>
        const byte _cc_EOT = 0x04;
        /// <summary>
        /// enquiry
        /// </summary>
        const byte _cc_ENQ = 0x05;
        /// <summary>
        /// acknowledge
        /// </summary>
        const byte _cc_ACK = 0x06;
        /// <summary>
        /// bell
        /// </summary>
        const byte _cc_BEL = 0x07;
        /// <summary>
        /// backspace
        /// </summary>
        const byte _cc_BS = 0x08;
        /// <summary>
        /// horizontal Tab
        /// </summary>
        const byte _cc_HT = 0x09;
        /// <summary>
        /// line feed
        /// </summary>
        const byte _cc_LF = 0x0a;
        /// <summary>
        /// vertical tab
        /// </summary>
        const byte _cc_VT = 0x0b;
        /// <summary>
        /// form feed
        /// </summary>
        const byte _cc_FF = 0x0c;
        /// <summary>
        /// carriage return
        /// </summary>
        const byte _cc_CR = 0x0d;
        /// <summary>
        /// shift out
        /// </summary>
        const byte _cc_SO = 0x0e;
        /// <summary>
        /// shift in
        /// </summary>
        const byte _cc_SI = 0x0f;
        /// <summary>
        /// data link escape
        /// </summary>
        const byte _cc_DLE = 0x10;
        /// <summary>
        /// device control 1
        /// </summary>
        const byte _cc_DC1 = 0x11;
        /// <summary>
        /// device control 2
        /// </summary>
        const byte _cc_DC2 = 0x12;
        /// <summary>
        /// device control 3
        /// </summary>
        const byte _cc_DC3 = 0x13;
        /// <summary>
        /// device control 4
        /// </summary>
        const byte _cc_DC4 = 0x14;
        /// <summary>
        /// negative acknowledge
        /// </summary>
        const byte _cc_NAK = 0x15;
        /// <summary>
        /// synchronous idle
        /// </summary>
        const byte _cc_SYN = 0x16;
        /// <summary>
        /// end of transmission block
        /// </summary>
        const byte _cc_ETB = 0x17;
        /// <summary>
        /// cancel
        /// </summary>
        const byte _cc_CAN = 0x18;
        /// <summary>
        /// end of medium
        /// </summary>
        const byte _cc_EM = 0x19;
        /// <summary>
        /// substitute
        /// </summary>
        const byte _cc_SUB = 0x1a;
        /// <summary>
        /// escape
        /// </summary>
        const byte _cc_ESC = 0x1b;
        /// <summary>
        /// file separator
        /// </summary>
        const byte _cc_FS = 0x1c;
        /// <summary>
        /// group separator
        /// </summary>
        const byte _cc_GS = 0x1d;
        /// <summary>
        /// record separator
        /// </summary>
        const byte _cc_RS = 0x1e;
        /// <summary>
        /// unit separator
        /// </summary>
        const byte _cc_US = 0x1f;
        /// <summary>
        /// rubout / delete
        /// </summary>
        const byte _cc_DEL = 0x7f;
        #endregion

        #region Member Variables
        private SerialCom _serial; //for serial communication
        #endregion

        #region Constructor / Destructor
        /// <summary>
        /// Create a new communication over the serial port
        /// </summary>
        /// <param name="serialPort">serial port which connected to the device</param>
        /// <param name="baudRate">baudrate of this connection</param>
        /// <param name="timeOutValue">value of the time until the timeout</param>
        public QCommunication(string serialPort, BaudRate baudRate, int timeOutValue)
        {
            _serial = new SerialCom();
            _serial.PortName = serialPort;
            _serial.Baudrate = getBaudRate(baudRate);
            _serial.Parity = SerialCom.Parities.None;
            _serial.DataBits = 8;
            _serial.Stopbit = SerialCom.StopBit.One;
            _serial.DTREnable = true;
            _serial.RTSEnable = true;
            _serial.TimeOutValue = timeOutValue;

            _serial.TimeOut += new EventHandler(_serial_TimeOut);
            _serial.DataReceived += new SerialCom.ReadingRecordEventHandler(_serial_DataReceived);
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~QCommunication()
        {
            Close();
        }
        #endregion

        #region Public Functions
        /// <summary>
        /// Close the QCommunication Session - include Dispose()
        /// </summary>
        public void Close()
        {
            Dispose();
        }

        /// <summary>
        /// IDisposable Function - release the ressources
        /// </summary>
        public void Dispose()
        {
            _serial.Dispose();
        }

        /// <summary>
        /// send a string command
        /// </summary>
        /// <param name="command">string command</param>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        public void SendCommand(string command, bool waitForAnswer)
        {
            sendOverSerial(command, waitForAnswer);
        }

        /// <summary>
        /// send a byte array command
        /// </summary>
        /// <param name="command">byte array command</param>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        public void SendCommand(byte[] command, bool waitForAnswer)
        {

                sendOverSerial(command, waitForAnswer);
        }

        /// <summary>
        /// send a acknowledge signal
        /// </summary>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        public void sendACK(bool waitForAnswer)
        {
           sendOverSerial_ACK(waitForAnswer);
        }

        /// <summary>
        /// send a not acknowlege signal
        /// </summary>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        public void sendNAK(bool waitForAnswer)
        {
                sendOverSerial_NAK(waitForAnswer);
 
        }

        /// <summary>
        /// reset the timeout value, it is important because the serial port check in intervall if a signal is coming
        /// is the reset time value old, so will the timeout signal trigger every time 
        /// </summary>
        public void resetSerialTimeoutValue()
        {
            if (_serial != null)
                _serial.resetTimeOutValue();
        }
        #endregion

        #region Events
        /// <summary>
        /// trigger by incoming data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _serial_DataReceived(object sender, ReceivedDataEventArgs e)
        {
            if (e.ReceivedData.Length > 1)
            {
                byte[] answer = e.ReceivedData;
                string sAnswer = string.Empty;
                byte LRC = 0x00;
                byte checkLRC = 0x00;

                bool beginMessage = false;
                bool endMessage = false;
                for (int i = 0; i < answer.Length; i++)
                {
                    if (answer[i] == _cc_STX)
                    {
                        LRC = 0x00;
                        beginMessage = true;
                        endMessage = false;
                        LRC += answer[i];
                    }
                    else if (beginMessage == true)
                    {
                        if (answer[i] == _cc_ETX)
                        {
                            beginMessage = false;
                            endMessage = true;
                        }
                        else
                        {
                            sAnswer += CommunicationTools.byteToString(answer[i]);
                        }
                        LRC += answer[i];
                    }
                    else if (endMessage)
                    {
                        checkLRC = answer[i];
                        break;
                    }
                }

                LRC = (byte)((0xff - LRC) + 0x01);

                if (LRC != checkLRC)
                {
                    if (OnCRCError != null)
                        //OnCRCError(this, new EventArgs());
                        ThreadSafe.Invoke(this.OnCRCError, new object[] { this, new QCommunicationCRCEventArgs(sAnswer, answer, checkLRC, LRC) });
                }

                if (OnIncomingData != null) //
                {
                    //OnIncomingData(this, new QCommunicationEventArgs(answer, decompressData));
                    ThreadSafe.Invoke(this.OnIncomingData, new object[] { this, new QCommunicationEventArgs(sAnswer, answer) });
                }
            }
            else
            {
                if (e.ReceivedData[0] == _cc_ACK)
                {
                    if (OnACKMessage != null)
                        ThreadSafe.Invoke(this.OnACKMessage, new object[] { this, EventArgs.Empty });
                }
                else if (e.ReceivedData[0] == _cc_NAK)
                {
                    if (OnNAKMessage != null)
                        ThreadSafe.Invoke(this.OnNAKMessage, new object[] { this, EventArgs.Empty });
                }
                else if (e.ReceivedData[0] == _cc_SYN)
                {
                    if (OnSYNMessage != null)
                        ThreadSafe.Invoke(this.OnSYNMessage, new object[] { this, EventArgs.Empty });
                }
            }

        }

        /// <summary>
        /// trigger by a timeout signal for the serial commonication
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _serial_TimeOut(object sender, EventArgs e)
        {
            if (OnTimeOut != null)
                //OnTimeOut(this, new EventArgs());
                ThreadSafe.Invoke(this.OnTimeOut, new object[] { this, EventArgs.Empty });
        }
        #endregion

        #region private Functions
		/// <summary>
        /// split the messages
        /// </summary>
        /// <param name="messages"></param>
        /// <returns></returns>
        private ArrayList getMessages(byte[] messages)
        {
            ArrayList mesList = new ArrayList();
            bool beginMessage = false;
            string sMessage = string.Empty;
            for (int i = 0; i < messages.Length - 2; i++)
            {
                if (beginMessage)
                    sMessage = sMessage + CommunicationTools.byteToString(messages[i]);
                
                if (messages[i] == _cc_STX)
                {
                    beginMessage = true;
                    sMessage = string.Empty;
                    sMessage = sMessage + CommunicationTools.byteToString(messages[i]) + CommunicationTools.byteToString(messages[i + 1]);
                    i++;
                }


                if (messages[i] == _cc_ETX && beginMessage)
                {
                    beginMessage = false;
                    sMessage = sMessage + CommunicationTools.byteToString(messages[i + 1]) + CommunicationTools.byteToString(messages[i + 2]);
                    mesList.Add(CommunicationTools.stringToByteArray(sMessage));
                    sMessage = string.Empty;
                    i += 2;
                }
            }

            return mesList;
        }
		
        /// <summary>
        /// send a string over serial port
        /// </summary>
        /// <param name="command">command as a string</param>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        private void sendOverSerial(string command, bool waitForAnswer)
        {
            byte[] bCommand = CommunicationTools.stringToByteArray(command);
            byte[] bSendCommand = createSendData(bCommand);
            _serial.sendData(bSendCommand, waitForAnswer);
        }

        /// <summary>
        /// send a byte array over serial port
        /// </summary>
        /// <param name="command">command as a byte array</param>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        private void sendOverSerial(byte[] command, bool waitForAnswer)
        {
            byte[] bSendCommand = createSendData(command);
            _serial.sendData(bSendCommand, waitForAnswer);
        }

        /// <summary>
        /// send the acknowledge information over serial port
        /// </summary>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        private void sendOverSerial_ACK(bool waitForAnswer)
        {
            byte[] bSendCommand = new byte[] {_cc_ACK};
            _serial.sendData(bSendCommand, waitForAnswer);
        }

        /// <summary>
        /// send a not acknowledge information over serial port
        /// </summary>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        private void sendOverSerial_NAK(bool waitForAnswer)
        {
            byte[] bSendCommand = new byte[] { _cc_NAK };
            _serial.sendData(bSendCommand, waitForAnswer);
        }

        /// <summary>
        /// get the Baudrate for the Serial port by commited baudrate for the boxen
        /// </summary>
        /// <param name="baudRate">baudrate for the boxen</param>
        /// <returns>serial port baudrate</returns>
        private SerialCom.BaudRate getBaudRate(BaudRate baudRate)
        {
            switch (baudRate)
            {
                case BaudRate.CBR_9600:               
                    return SerialCom.BaudRate.CBR_9600;
                case BaudRate.CBR_19200:
                    return SerialCom.BaudRate.CBR_19200;
                case BaudRate.CBR_38400:
                    return SerialCom.BaudRate.CBR_38400;
                case BaudRate.CBR_57600:
                    return SerialCom.BaudRate.CBR_57600;
                case BaudRate.CBR_115200:
                    return SerialCom.BaudRate.CBR_115200;
                default:
                    return SerialCom.BaudRate.CBR_9600;
            }
        }

        /// <summary>
        /// created the full send data with header, data, control character and LRC value
        /// </summary>
        /// <param name="sendText">data text will be sent</param>
        /// <returns>full created send bytes</returns>
        private byte[] createSendData(byte[] sendText)
        {
            byte[] bSendData = new byte[sendText.Length + 2];

            bSendData[0] = _cc_STX;
            int i = 0;

            for (; i < sendText.Length; i++)
            {
                bSendData[i + 1] = sendText[i];
            }

            i++;

            bSendData[i] = _cc_ETX;

            byte LRC = 0;

            for (int a = 0; a < bSendData.Length; a++)
            {
                LRC += bSendData[a];
            }

            LRC = (byte)((0xff - LRC) + 0x01);

            byte[] bSendData2 = new byte[bSendData.Length + 1];

            i = 0;
            for (; i < bSendData.Length; i++)
            {
                bSendData2[i] = bSendData[i];
            }

            bSendData2[i] = LRC;
            

            return bSendData2;
        }
        #endregion

        #region Compression
        /// <summary>
        /// decompress a dataarray - 0x00 and 0xff values are compress after crc-calculation
        /// -before computing data, the dataarray musst decompress
        /// 0x10 in data comes twice it is one
        /// </summary>
        /// <param name="data">dataarray wich will be decompress</param>
        /// <returns>decompress data array</returns>
        private byte[] decompressDataStream(byte[] data)
        {
            if (data.Length > 6)
            {
				int length = data.Length;
				string sDecomp = "";
				int posLastDLE = int.MinValue;
				//DLE ETX
                sDecomp = sDecomp + CommunicationTools.byteToString(data[0]);
                sDecomp = sDecomp + CommunicationTools.byteToString(data[1]);

				//Header and Data decompress
				for (int i = 2; i < length - 4; i++)
				{
					if (data[i] == 0x10)
					{
						if ((posLastDLE + 1) == i)
						{
							i++;
						}
						else
						{
							posLastDLE = i;
						}
					}

					if ((data[i] == 0x00 || data[i] == 0xff) && (i + 1 != length))
					{
						int count = (int)data[i + 1];
						for (int j = 0; j < count; j++)
						{
							sDecomp = sDecomp + CommunicationTools.byteToString(data[i]);
						}
						i++;
					}
					else
					{
						sDecomp = sDecomp + CommunicationTools.byteToString(data[i]);
					}
				}

				//DLE STX and CRC
                for (int i = length - 4; i < length; i++)
                {
                    sDecomp = sDecomp + CommunicationTools.byteToString(data[i]);
                }
				
				return CommunicationTools.stringToByteArray(sDecomp);
			}
			else
				return data;
		}

        /// <summary>
        /// compress a data array - 0x00 and 0xff values will be compress 
        /// 0x10 must send in twice in the data set
        /// </summary>
        /// <param name="data">uncompress data-array</param>
        /// <returns>compress data-array</returns>
        private byte[] compressDataStream(byte[] data)
        {
            int length = data.Length;
            string sComp = "";
            byte compZeroCounter = 0;
            byte compFFCounter = 0;

            for (int i = 0; i < length; i++)
            {
                
				if (data[i] == 0x10)
                {
                    sComp = sComp + CommunicationTools.byteToString(0x10) + CommunicationTools.byteToString(0x10);
                }
				
				if (data[i] == 0x00 && compFFCounter == 0)
                {
                    compZeroCounter++;
                }
                else if (data[i] == 0xff && compZeroCounter == 0)
                {
                    compFFCounter++;
                }
                else
                {
                    if (compFFCounter != 0)
                    {
                        sComp = sComp + CommunicationTools.byteToString(0xff) + CommunicationTools.byteToString(compFFCounter);
                        compFFCounter = 0;

                        if (data[i] == 0x00)
                            compZeroCounter++;
                    }
                    else if (compZeroCounter != 0)
                    {
                        sComp = sComp + CommunicationTools.byteToString(0x00) + CommunicationTools.byteToString(compZeroCounter);
                        compZeroCounter = 0;

                        if (data[i] == 0xff)
                            compFFCounter++;
                    }

                    if (compZeroCounter == 0 && compFFCounter == 0)
                    {
                        if (data[i] == 0x10)
                        {
                            sComp = sComp + CommunicationTools.byteToString(data[i]) + CommunicationTools.byteToString(data[i]);
                        }
                        else
                        {
                            sComp = sComp + CommunicationTools.byteToString(data[i]);
                        }
                    }
                    else if (i == length - 1)
                    {
                        if (compFFCounter != 0)
                        {
                            sComp = sComp + CommunicationTools.byteToString(0xff) + CommunicationTools.byteToString(compFFCounter);
                        }
                        else if (compZeroCounter != 0)
                        {
                            sComp = sComp + CommunicationTools.byteToString(0x00) + CommunicationTools.byteToString(compZeroCounter);
                        }
                    }
                }
            }

            return CommunicationTools.stringToByteArray(sComp);
            //return CommunicationTools.unicodeStringToASCIIByteArray(sComp, codePage);
        }
        #endregion

        #region CRC Check
        /// <summary>
        /// check the crc value of the data
        /// </summary>
        /// <param name="data">incoming data array</param>
        /// <returns>is crc value correct</returns>      
        private bool isCRCOK(byte[] data)
        {
            try
            {
                int length = data.Length;
                byte[] crcCheckSum = new byte[] { data[length - 2], data[length - 1] };

                byte[] rData = new byte[length - 6];

                for (int i = 2; i < length - 4; i++)
                {
                    rData[i - 2] = data[i];
                }

                byte[] calcCRC = CalculateCRC(rData);

                if (calcCRC[0] == crcCheckSum[0])
                {
                    if (calcCRC[1] == crcCheckSum[1])
                    {
                        return true;
                    }
                    else
                        return false;
                }
                else
                    return false;
            }
            catch
            {
                return false;
            }

        }

        /// <summary>
        /// CRC Lookup-Table
        /// </summary>
        UInt32[] CRC_Table = new UInt32[] 
            {
            0x0000, 0xCC01,	0xD801,	0x1400,	0xF001,	0x3C00,	0x2800,	0xE401,
	        0xA001,	0x6C00,	0x7800,	0xB401,	0x5000,	0x9C01,	0x8801,	0x4400
            };

        /// <summary>
        /// Calculate for the commited data-array a 16-bit crc value
        /// </summary>
        /// <param name="data">data array</param>
        /// <returns>16 bit crc value</returns>
        private byte[] CalculateCRC(byte[] data)
        {
            UInt32 crc = 0xffff;

            foreach (byte b in data)
            {
                crc = UpdateCRC(crc, b);
            }

            byte[] bCRC = BitConverter.GetBytes(crc);
            byte[] bCRC16 = new byte[2];

            //Filter 16 Bit
            if (bCRC.Length == 4)
            {
                for (int i = 0; i < 2; i++)
                {
                    bCRC16[i] = bCRC[1 - i];
                }
            }
            else
                bCRC16 = bCRC;

            return bCRC16;
        }

        /// <summary>
        /// Update the old crc value with the new calculated value
        /// </summary>
        /// <param name="crc">old crc value</param>
        /// <param name="data">data byte, which for the new crc-value needed</param>
        /// <returns>updated crc value</returns>
        private UInt32 UpdateCRC(UInt32 crc, byte data)
        {
            UInt32 ch = data;

            crc = CRC_Table[(ch ^ crc) & 15] ^ (crc >> 4);
            crc = CRC_Table[((ch >> 4) ^ crc) & 15] ^ (crc >> 4);

            return crc;
        }
        #endregion
    }

    /// <summary>
    /// Eventclass for QCommunication
    /// </summary>
    public class QCommunicationEventArgs : EventArgs
    {
        string _receiveData;
        byte[] _receiveDataAsByteArray;

        /// <summary>
        /// incoming Data as a String 
        /// only the informations without control characters
        /// </summary>
        public string ReceiveMessageString
        {
            get { return this._receiveData; }
        }

        /// <summary>
        /// the complete incoming data as a byte array
        /// </summary>
        public byte[] ReceiveMessageByte
        {
            get { return this._receiveDataAsByteArray; }
        }

        /// <summary>
        /// Evenarguments constructor 
        /// </summary>
        /// <param name="sMessage">incoming data as a string</param>
        /// <param name="bMessage">incoming data as a byte array</param>
        public QCommunicationEventArgs(string sMessage, byte[] bMessage)
        {
            this._receiveData = sMessage;
            this._receiveDataAsByteArray = bMessage;
        }
    }

    /// <summary>
    /// Eventclass for the CRC Error Event
    /// </summary>
    public class QCommunicationCRCEventArgs : EventArgs
    {
        string _receiveData;
        byte[] _receiveDataAsByteArray;
        byte _incomCRCValue;
        byte _exceptCRCValue;

        /// <summary>
        /// incoming Data as a String 
        /// only the informations without control characters
        /// </summary>
        public string ReceiveMessageString
        {
            get { return this._receiveData; }
        }

        /// <summary>
        /// the complete incoming data as a byte array
        /// </summary>
        public byte[] ReceiveMessageByte
        {
            get { return this._receiveDataAsByteArray; }
        }

        /// <summary>
        /// the value of the receive CRC
        /// </summary>
        public byte IncomingCRCValue
        {
            get { return this._incomCRCValue; }
        }

        /// <summary>
        ///  the value of the except CRC (calculated)
        /// </summary>
        public byte ExceptCRCValue
        {
            get { return this._exceptCRCValue; }
        }

        /// <summary>
        /// Eventargument Constructor
        /// </summary>
        /// <param name="sMessage">incoming data as a string</param>
        /// <param name="bMessage">incoming data as a byte array</param>
        /// <param name="incomCRC">incoming crc value</param>
        /// <param name="exceptCRC">calculated crc value</param>
        public QCommunicationCRCEventArgs(string sMessage, byte[] bMessage, byte incomCRC, byte exceptCRC)
        {
            this._receiveData = sMessage;
            this._receiveDataAsByteArray = bMessage;
            this._incomCRCValue = incomCRC;
            this._exceptCRCValue = exceptCRC;
        }
    }

    /// <summary>
    /// Helper class for thread safety event calling
    /// </summary>
    static class ThreadSafe
    {
        /// <summary>
        /// Method call event event safety - for using event beyond threads 
        /// </summary>
        /// <param name="method">event delegate</param>
        /// <param name="args">event arguments - the first one is the calling method like "this"</param>
        public static void Invoke(Delegate method, object[] args)
        {
            if (method != null)
            {
                foreach (Delegate handler in method.GetInvocationList())
                {
                    if (handler.Target is System.Windows.Forms.Control)
                    {
                        System.Windows.Forms.Control target = handler.Target as System.Windows.Forms.Control;

                        if (target.IsHandleCreated)
                        {
                            target.BeginInvoke(handler, args);
                        }
                    }
                    else if (handler.Target is System.ComponentModel.ISynchronizeInvoke)
                    {
                        System.ComponentModel.ISynchronizeInvoke target = handler.Target as System.ComponentModel.ISynchronizeInvoke;
                        target.BeginInvoke(handler, args);
                    }
                    else
                    {
                        handler.DynamicInvoke(args);
                    }
                }
            }
        }
    }
}
