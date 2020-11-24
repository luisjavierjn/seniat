/*******************************************************************************
$Id: //depot/Projects_Win_Rico/Seniat_net/Main/src/Seniat/Seniat/SerialCom.cs#4 $
$DateTime: 2011/07/28 14:07:18 $
$Change: 721 $
$Revision: #4 $
$Author: Zabel $
Copyright: ©2011 QUORiON Data Systems GmbH
*******************************************************************************/
using System;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Collections;

namespace QCom
{
    /// <summary>
    /// Class for control the Serial Communication
    /// Encapsulated the System.IO.Ports.SerialPorts
    /// </summary>
    class SerialCom
    {

        #region Structs/Enumerations
        /// <summary>
        /// Enumerate the possible Baudrates
        /// </summary>
        public enum BaudRate
        {
            CBR_1200,
            CBR_2400,
            CBR_4800,
            CBR_9600,
            CBR_19200,
            CBR_38400,
            CBR_57600,
            CBR_115200
        }

        /// <summary>
        /// Enumerate the possible Stopbits
        /// </summary>
        public enum StopBit
        {
            One,
            OnePointFive,
            Two
        }

        /// <summary>
        /// Enumerate the possible Parities
        /// </summary>
        public enum Parities
        {
            Even,
            Mark,
            None,
            Odd,
            Space
        }

        /*
        /// <summary>
        /// Enumerate the predefinied device configurations
        /// </summary>
        public enum DeviceConfiguration
        {
            BOXEN,
            CR30
        }*/
        #endregion

        #region EventHandler
        /// <summary>
        /// Event Handler for the serial data receive indication
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public delegate void ReadingRecordEventHandler(object sender, ReceivedDataEventArgs e);
        /// <summary>
        /// Event indicated by incoming data
        /// </summary>
        public event ReadingRecordEventHandler DataReceived;
        /// <summary>
        /// Evend indicated by timeout - no response from the source
        /// </summary>
        public event EventHandler TimeOut;
        #endregion

        #region Member Variables
        private System.IO.Ports.SerialPort _sPort; //Serial networkPort
        private BackgroundWorker _readThread = null;
        private int _timeOut = 5000; //time variable
        private static bool _continue = false;
        private long _timeOutStartTicks = 0;
        private int _waitingTime = 200;

#if DEBUG
        private string logFilename = "serialLOG.log";
#endif

        #endregion

        #region Encapsulation
        /// <summary>
        /// Available Ports for this Computer
        /// </summary>
        public string[] PortNames
        {
            get
            {
                try
                {
                    return System.IO.Ports.SerialPort.GetPortNames();
                }
                catch
                {
                    return new string[0];
                }
            }
        }

        /// <summary>
        /// get or set the baudrate for the communication
        /// if setting the baudrate by active connection, the connection will be closed
        /// </summary>
        public BaudRate Baudrate
        {
            get { return getBaudRate(); }
            set
            {
                setBaudRate(value);
            }
        }

        /// <summary>
        /// get or set the PortName
        /// if setting the portname by active connection, the connection will be closed
        /// the available Portname can be calling by the varialbe PortNames
        /// </summary>
        public string PortName
        {
            get { return this._sPort.PortName; }
            set
            {
                setPortName(value);
            }
        }

        /// <summary>
        /// get or set the length of data bits per bytes for the connection
        /// if setting by open connection, the connection will be closed
        /// </summary>
        public int DataBits
        {
            get { return this._sPort.DataBits; }
            set { setDatabits(value); }
        }

        /// <summary>
        /// get or set the parity-checking protocol
        /// if setting by open connection, the connection will be closed
        /// </summary>
        public Parities Parity
        {
            get { return getParity(); }
            set
            {
                setParity(value);
            }
        }

        /// <summary>
        /// get and set the number of stopbits per byte
        /// if setting by open connection, the connection will be closed
        /// </summary>
        public StopBit Stopbit
        {
            get { return getStopbit(); }
            set { setStopbit(value); }
        }

        /// <summary>
        /// get or set a value that enables the Data Terminal Ready (DTR) signal during the serial communication
        /// if setting by open connection, the connection will be closed
        /// </summary>
        public bool DTREnable
        {
            get { return this._sPort.DtrEnable; }
            set { setDtrEnable(value); }
        }

        /// <summary>
        /// get or set a value indicating whether the Request to Send (RTS) signal is enabled during serial communication
        /// if setting by active connection, the connection will be closed
        /// </summary>
        public bool RTSEnable
        {
            get { return this._sPort.RtsEnable; }
            set { setRtsEnable(value); }
        }

        public int TimeOutValue
        {
            get { return this._timeOut; }
            set 
            { 
                    
                this._timeOut = value;
                if (_sPort != null)
                {
                    _sPort.ReadTimeout = this._timeOut;
                    _sPort.WriteTimeout = this._timeOut;
                }
            }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Constructor
        /// </summary>
        public SerialCom()
        {
            initSerial();
        }

        /// <summary>
        /// Constructor for initialisation of this class
        /// </summary>
        /// <param name="portName">Name of the networkPort, where connected the device</param>
        /// <param name="baudRate">baudrate for the device-communicaton</param>
        /// <param name="databits">count of the databits for the device communication</param>
        /// <param name="parity">kind of parity for the device communication</param>
        /// <param name="stopBit">kind of stopbits for the device communication</param>
        public SerialCom(string portName, BaudRate baudRate, int databits, Parities parity, StopBit stopBit)
        {
            initSerial();
            setPortName(portName);
            setBaudRate(baudRate);
            setDatabits(databits);
            setParity(parity);
            setStopbit(stopBit);
        }

        #endregion

        #region Public
        public void Dispose()
        {
            if (_readThread != null)
            {
                if (_readThread.IsBusy)
                {
                    _readThread.Dispose();
                    _readThread = null;
                }
            }
            
            _sPort.Close();
            _sPort.Dispose();
        }

        /// <summary>
        /// open a connection to the device
        /// </summary>
        public void openSerialPort()
        {
            if (_sPort.IsOpen)
                _sPort.Close();

            try
            {
                _sPort.Open();
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                    throw new Exception(e.InnerException.Message);
                else
                    throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// close the connection to the device
        /// </summary>
        public void closeSerialPort()
        {
            try
            {
                _sPort.Close();
            }
            catch (Exception e)
            {

                if (e.InnerException != null)
                    throw new Exception(e.InnerException.Message);
                else
                    throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// sending a string over the serial port
        /// </summary>
        /// <param name="text">sending text</param>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        public void sendText(string text, bool waitForAnswer)
        {
            if (!_sPort.IsOpen)
                openSerialPort();

            if(waitForAnswer)
                startReadThread();
#if DEBUG
            writeSerialLog(CommunicationTools.stringToByteArray(text), "SEND TO DEVICE:");
#endif
            _sPort.WriteLine(text);
        }

        /// <summary>
        /// sending a data byte array over the serial port
        /// </summary>
        /// <param name="data">sendign byte array</param>
        /// <param name="waitForAnswer">is set false the program will not wait for an answer</param>
        public void sendData(byte[] data, bool waitForAnswer)
        {
            if (!_sPort.IsOpen)
                openSerialPort();

            if(waitForAnswer)
                startReadThread();
#if DEBUG
            writeSerialLog(data, "SEND TO DEVICE:");
#endif
            _sPort.Write(data, 0, data.Length);
        }

        public void resetTimeOutValue()
        {
            this._timeOutStartTicks = DateTime.Now.Ticks;
        }
        #endregion

        #region nonPublic
        /// <summary>
        /// initialisation of the serial communication - sets events and standard configurations
        /// </summary>
        private void initSerial()
        {            
            _sPort = new System.IO.Ports.SerialPort();
            _sPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(_sPort_DataReceived);
            //_sPort.ErrorReceived += new System.IO.Ports.SerialErrorReceivedEventHandler(_sPort_ErrorReceived);
            _sPort.WriteTimeout = _timeOut;
            _sPort.ReadTimeout = _timeOut;

            _sPort.DtrEnable = true;
            _sPort.RtsEnable = true;
            _sPort.DiscardNull = false;

        }

        /// <summary>
        /// set the baudrate for the serial communication
        /// </summary>
        /// <param name="baudRate">constant baudrate for the communication</param>
        private void setBaudRate(BaudRate baudRate)
        {
            if (_sPort.IsOpen)
                _sPort.Close();
            try
            {
                switch (baudRate)
                {
                    case BaudRate.CBR_1200:
                        _sPort.BaudRate = 1200;
                        _waitingTime = 1400;
                        break;
                    case BaudRate.CBR_2400:
                        _sPort.BaudRate = 2400;
                        _waitingTime = 1000;
                        break;
                    case BaudRate.CBR_4800:
                        _sPort.BaudRate = 4800;
                        _waitingTime = 900;
                        break;
                    case BaudRate.CBR_9600:
                        _sPort.BaudRate = 9600;
                        _waitingTime = 700;
                        break;
                    case BaudRate.CBR_19200:
                        _sPort.BaudRate = 19200;
                        _waitingTime = 550;
                        break;
                    case BaudRate.CBR_38400:
                        _sPort.BaudRate = 38400;
                        _waitingTime = 400;
                        break;
                    case BaudRate.CBR_57600:
                        _sPort.BaudRate = 57600;
                        _waitingTime = 200;
                        break;
                    case BaudRate.CBR_115200:
                        _sPort.BaudRate = 115200;
                        _waitingTime = 100;
                        break;
                }
            }
            catch (ArgumentOutOfRangeException e)
            {
                throw new ArgumentOutOfRangeException(e.Message);
            }
            catch (IOException e)
            {
                throw new IOException(e.Message);
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                    throw new Exception(e.InnerException.Message);
                else
                    throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// return the setting baudrate
        /// </summary>
        /// <returns>baudrate is setting</returns>
        private BaudRate getBaudRate()
        {
            int baudRate = _sPort.BaudRate;
            switch (baudRate)
            {
                case 1200:
                    return BaudRate.CBR_1200;
                case 2400:
                    return BaudRate.CBR_2400;
                case 4800:
                    return BaudRate.CBR_4800;
                case 9600:
                    return BaudRate.CBR_9600;
                case 19200:
                    return BaudRate.CBR_19200;
                case 38400:
                    return BaudRate.CBR_38400;
                case 57600:
                    return BaudRate.CBR_57600;
                case 115200:
                    return BaudRate.CBR_115200;
                default:
                    return BaudRate.CBR_9600;
            }
        }

        /// <summary>
        /// set the port for the communication
        /// </summary>
        /// <param name="portname">portname for the communication</param>
        private void setPortName(string portname)
        {
            if (_sPort.IsOpen)
                _sPort.Close();
            if (isValidPortName(portname))
                _sPort.PortName = portname;
            else
                throw new ArgumentException("The portname " + portname + " is not guilty or don´t exist!");
        }

        /// <summary>
        /// check if the portname a valid or existing port on this device
        /// </summary>
        /// <param name="portname">portname which will check</param>
        /// <returns>is port correct</returns>
        private bool isValidPortName(string portname)
        {
            try
            {
                foreach (string pName in this.PortNames)
                {
                    string tName = pName.ToLower();
                    portname = portname.ToLower();

                    if (portname == tName)
                        return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// set the count of the databits for the communication
        /// </summary>
        /// <param name="databits">count of databits</param>
        private void setDatabits(int databits)
        {
            if (_sPort.IsOpen)
                _sPort.Close();

            if (databits < 5 || databits > 8)
            {
                throw new ArgumentOutOfRangeException("The Databits must be between 5 and 8 bits");
            }
            else
                _sPort.DataBits = databits;


        }

        /// <summary>
        /// set the kind of parity for the communication
        /// </summary>
        /// <param name="parity">parity</param>
        private void setParity(Parities parity)
        {
            try
            {
                if (_sPort.IsOpen)
                    _sPort.Close();
                switch (parity)
                {
                    case Parities.Even:
                        _sPort.Parity = System.IO.Ports.Parity.Even;
                        break;
                    case Parities.Mark:
                        _sPort.Parity = System.IO.Ports.Parity.Mark;
                        break;
                    case Parities.None:
                        _sPort.Parity = System.IO.Ports.Parity.None;
                        break;
                    case Parities.Odd:
                        _sPort.Parity = System.IO.Ports.Parity.Odd;
                        break;
                    case Parities.Space:
                        _sPort.Parity = System.IO.Ports.Parity.Space;
                        break;
                }
            }
            catch (ArgumentOutOfRangeException e)
            {
                throw new ArgumentOutOfRangeException(e.Message);
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                    throw new Exception(e.InnerException.Message);
                else
                    throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// return the setting parity
        /// </summary>
        /// <returns>partity wich is set</returns>
        private Parities getParity()
        {
            switch (_sPort.Parity)
            {
                case System.IO.Ports.Parity.Even:
                    return Parities.Even;
                case System.IO.Ports.Parity.Mark:
                    return Parities.Mark;
                case System.IO.Ports.Parity.None:
                    return Parities.None;
                case System.IO.Ports.Parity.Odd:
                    return Parities.Odd;
                case System.IO.Ports.Parity.Space:
                    return Parities.Space;
                default:
                    return Parities.None;
            }
        }

        /// <summary>
        /// set the amount of the stopbits for the serial communication
        /// </summary>
        /// <param name="stopbits">stopbit which set for the communication</param>
        private void setStopbit(StopBit stopbits)
        {
            try
            {
                if (_sPort.IsOpen)
                    _sPort.Close();

                switch (stopbits)
                {
                    case StopBit.One:
                        this._sPort.StopBits = System.IO.Ports.StopBits.One;
                        break;
                    case StopBit.OnePointFive:
                        this._sPort.StopBits = System.IO.Ports.StopBits.OnePointFive;
                        break;
                    case StopBit.Two:
                        this._sPort.StopBits = System.IO.Ports.StopBits.Two;
                        break;
                }
            }
            catch (ArgumentOutOfRangeException e)
            {
                throw new ArgumentOutOfRangeException(e.Message);
            }
            catch (Exception e)
            {
                if (e.InnerException != null)
                    throw new Exception(e.InnerException.Message);
                else
                    throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// return the setting stopbits for the communication
        /// </summary>
        /// <returns>setting stopbit</returns>
        private StopBit getStopbit()
        {
            switch (this._sPort.StopBits)
            {
                case System.IO.Ports.StopBits.One:
                    return StopBit.One;
                case System.IO.Ports.StopBits.OnePointFive:
                    return StopBit.OnePointFive;
                case System.IO.Ports.StopBits.Two:
                    return StopBit.Two;
                default:
                    return StopBit.One;
            }
        }

        /// <summary>
        /// set a value thats enable the data terminal ready signal during the communication
        /// </summary>
        /// <param name="dtrEnable">dis-/enable DTR</param>
        private void setDtrEnable(bool dtrEnable)
        {
            try
            {
                if (_sPort.IsOpen)
                    _sPort.Close();
                _sPort.DtrEnable = dtrEnable;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// return the RTS setting
        /// </summary>
        /// <param name="rtsEnalbe">RTS Setting</param>
        private void setRtsEnable(bool rtsEnalbe)
        {
            try
            {
                if (_sPort.IsOpen)
                    _sPort.Close();
                _sPort.RtsEnable = rtsEnalbe;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// convert a byte-array to a string
        /// </summary>
        /// <param name="data">byte array</param>
        /// <returns>string</returns>
        private string byteArrayToString(byte[] data)
        {
            string sData = "";
            foreach (byte b in data)
            {
                sData = sData + Convert.ToChar(b);
            }
            return sData;
        }

#if DEBUG
        private void writeSerialLog(byte[] values, string title)
        {
            StreamWriter sw = new StreamWriter(this.logFilename, true, System.Text.Encoding.Default);
            try
            {
                sw.WriteLine(title);
                DateTime now = DateTime.Now;
                sw.WriteLine(now.ToLongDateString() + " " + now.ToLongTimeString() + "," + now.Millisecond.ToString("000"));
                sw.WriteLine(CommunicationTools.byteArrayToString(values));
                sw.WriteLine("HEX:");
                sw.WriteLine(CommunicationTools.ByteArrayToHexString(values));
                sw.WriteLine("LENGTH: " + values.Length.ToString());
                sw.WriteLine();
            }
            catch
            { }
            finally
            {
                sw.Close();
            }
        }
#endif
        #endregion

        #region EventHandling
        /// <summary>
        /// occurs when an error send from the io-ports - currently not implemented
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _sPort_ErrorReceived(object sender, System.IO.Ports.SerialErrorReceivedEventArgs e)
        {
            throw new NotImplementedException();

        }

        /// <summary>
        /// occure whent data received over the io port
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _sPort_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            /*
            if (incomingBytes == null)
                incomingBytes = new ArrayList();

            _continue = false;

            int length = _sPort.BytesToRead;
            byte[] readChars = new byte[length];

            _sPort.Read(readChars, 0, length);

            incomingBytes.Add(readChars);
            incomingBytesCount += length;

            for (int i = 0; i < readChars.Length; i++)
            {
                if (readChars[i] == 3)
                {
                    byte[] value = new byte[incomingBytesCount];
                    int counter = 0;

                    foreach (byte[] bytes in incomingBytes)
                    {
                        for (int j = 0; j < bytes.Length; j++)
                        {
                            value[counter] = bytes[j];
                        }
                    }

                    if (DataReceived != null)
                    {
                        string readText = byteArrayToString(value);
                        DataReceived(this, new ReceivedDataEventArgs(readText, value));
                    }

                    incomingBytes = null;
                    incomingBytesCount = 0;

                    break;
                }
            }*/


            
            try
            {
                Thread.Sleep(_waitingTime); //time for waiting until all data received in buffer
                int length = _sPort.BytesToRead;

                byte[] readChars = new byte[length];

                _sPort.Read(readChars, 0, length);

                if (readChars.Length > 0)
                {
                    _continue = false;
#if DEBUG
                    writeSerialLog(readChars, "INCOMING FROM DEVICE");
#endif
                    if (DataReceived != null)
                    {
                        string readText = byteArrayToString(readChars);
                        DataReceived(this, new ReceivedDataEventArgs(readText, readChars));
                    }
                }
            }
            catch
            {
                
            }
        }
        #endregion

        #region Thread
        private void startReadThread()
        {
            if (_readThread != null)
            {
                _readThread.Dispose();
                _readThread = null;
            }
            
            if (_readThread == null)
            {
                _readThread = new BackgroundWorker();
                _readThread.WorkerSupportsCancellation = true;
                _readThread.WorkerReportsProgress = true;
                _readThread.DoWork += new DoWorkEventHandler(readThread_DoWork);
                _readThread.RunWorkerCompleted += new RunWorkerCompletedEventHandler(readThread_RunWorkerCompleted);
            }

            
            _continue = true;
            _readThread.RunWorkerAsync();
        }

        private void stopReadThread()
        {
            _continue = false;

            if(_sPort != null)
                _sPort.Close();

            if (_readThread != null)
            {
                _readThread.Dispose();
                _readThread = null;
            }
        }

        void readThread_DoWork(object sender, DoWorkEventArgs e)
        {
            //long startTicks = DateTime.Now.Ticks;
            resetTimeOutValue();

            while (_continue)
            {
                long stopTicks = DateTime.Now.Ticks;
                long durance = stopTicks - _timeOutStartTicks;

                long miliSeconds = durance / TimeSpan.TicksPerMillisecond;

                if (miliSeconds > _timeOut)
                {
                    if (TimeOut != null)
                    {
                        _continue = false;
                        TimeOut(this, new EventArgs());

                    }
                }
            }
        }


        void readThread_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //throw new NotImplementedException();
            stopReadThread();
        }
        #endregion
    }

    /// <summary>
    /// class definies the event args for the incoming data 
    /// </summary>
    public class ReceivedDataEventArgs : EventArgs
    {
        string _receivedText;
        byte[] _receivedData;

        #region Encapsulation
        /// <summary>
        /// incoming data as string
        /// </summary>
        public string ReceivedText
        {
            get { return this._receivedText; }
        }

        /// <summary>
        /// incoming data as byte array
        /// </summary>
        public byte[] ReceivedData
        {
            get { return this._receivedData; }
        }
        #endregion

        #region Constructor
        /// <summary>
        /// constructor for this class
        /// </summary>
        /// <param name="receivedText">incoming serial data as string</param>
        /// <param name="receivedData">incoming serial data as byte-array</param>
        public ReceivedDataEventArgs(string receivedText, byte[] receivedData)
        {
            this._receivedData = receivedData;
            this._receivedText = receivedText;
        }
        #endregion
    }
}

