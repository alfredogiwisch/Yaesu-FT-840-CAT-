# Yaesu-FT-840-CAT-

Frontend application with installer package to control via RS-232 the Yaesu FT-840 transceiver

This application for the remote control and operation of the Yaesu FT-840 transceiver was initally developed in year 2001 with Visual Basic 6.0. 
It requires a serial port to TTL interface like the MAX232. The latest change was the adding of a bargraph meter to meassure the incomming RX 
signal and TX power.

After installation please select an available serial port. 

I don`t tested the application with a USB to serial cable but I think it would work because the drivers are PNP compatible with RS-232 standards.

In the files there is a schematoic of the TTL to RS232 interface using the MAX 232 chip.

For more technicla information information please visit my blog 

http://alfredoblogspage.blogspot.com/2007/05/yaesu-ft-840-cat-controller-code.html

How it works:

This tutorial covers the main code procedures used in the CAT controller app that I have developed under Visual Basic 6. The Yaesu FT-840 is a communication equipment that covers the Ham Radio HF bands (160, 80, 40, 20, 15, 10 meters). It use the RS232 serial port for connection to a computer. The equipment I/O serial port level is TTL +5V, however a TTL to RS232 converter is necessary because on a PC the levels are +-12V. 

A cheap option to build is the TTL to RS232 signals converter integrated circuit MAX232. The CAT interface is very easy to build so I strongly recommend to do it yourself.

The C.A.T. (Computer Aided Transceiver) protocol provides complete control from a PC. Operations such TX/RX mode selection, frequency input, memory storage and retrieve, transceiver status data dump and others functions are available. All the commands sent to the equipment consist of blocks of five bytes. The last byte sent in each block is the instruction opcode and previous four bytes the arguments. 
Below is an command block example that set the "AM wide reception" operation "MODE" on the transceiver:

MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(4)
MSComm1.Output = Chr$(12)

The first four bytes are the arguments. The parameter $4 select "AM wide reception" and the opcode $12 set the command "MODE" (TX/RX mode on AM procedure).
 
Before starting to send or receive data the serial port must be set to 4800 bauds, 1 start bit, 8 bit data, 2 stops bits, no parity. The following procedure initialize the serial port: 
MSComm1.CommPort = Text1
MSComm1.Settings = "4800,N,8,2"
MSComm1.PortOpen = True
MSComm1.InputLen = 5
 
The main procedures are the numeric key array for frequency operation entry and the numeric data conversions. The numeric keypad array use the index parameter to know which key is pressed. When the number of input digits reach 7 the procedure Sendfrec_Click is called and txtNumero variable is passed as parameter to evaluate first if the input value is OK or out of range. 
Private Sub cmdBotones_Click (index As Integer) Static a As Integer
MCI.From = 0
MCI.Command = "Play"
If txtNumero = "" Then a = 0
txtNumero = txtNumero + Format (index)
a = a + 1
If a = 7 Then
Sendfrec_Click
End If
End Sub
 
Before the numeric data conversion the program verify if the frequency value aka txtNumero parameter is out of range. The frequency coverage range on the FT-840 goes from 100.000 (100 khz) to 30.000.000 (30 mhz). Any value out of this range shows a warning message box to re enter the frequency. 
The following procedure check if the numeric value is out of range: 
Private Sub Sendfrec_Click()
Dim n1$, n2$, n3$, n4$
Dim cadena As Boolean
Dim z1, z2, z3, z4, pa1, pa2, pa3, pa4 As Integer
Static numero As String
On Error GoTo Manejoerror
cadena = txtNumero Like "#######"
If cadena = False Then
MsgBox "Please input frecuency value ", vbCritical, "Yaesu FT-840 Error"
txtNumero = numero
End If
If txtNumero = "" Then
MsgBox "Please input frecuency value ", vbCritical, "Yaesu FT-840 Error"
txtNumero = numero
Exit Sub
End If
If txtNumero >= 3000001 Then
MsgBox "The frecuency value must be < txtnumero = "" txtnumero =" numero"> 100.00 Khz", vbCritical, "Yaesu FT-840 Error"
txtNumero = ""
txtNumero = numero
End If
End If
 
The numeric data conversions below is the key procedure of the program to convert txtNumero variable to the CAT protocol format accepted by the trasnceiver. The procedure use numerical functions for example the val function to return the numeral part of a string. The use of the functions left$, mid$, right$ split the string chain because the parameters for frequency operation need to be separated into 2 digit blocks to build the four bytes argument required in the CAT protocol. The last five instructions send the argument result and the command instruction in CAT protocol format to the Yaesu FT-840 transceiver. 
numero = txtNumero
n1$ = Left$(txtNumero, 1)
n2$ = Mid$(txtNumero, 2, 2)
n3$ = Mid$(txtNumero, 4, 2)
n4$ = Right$(txtNumero, 2)
z1 = Val(n1$)
z2 = Val(n2$)
z3 = Val(n3$)
z4 = Val(n4$)
If z2 >= 10 Then
ze2 = Left(z2, 1)
z2 = z2 + (ze2 * 6)
End If
If z3 >= 10 Then
ze3 = Left(z3, 1)
z3 = z3 + (ze3 * 6)
End If
If z4 >= 10 Then
ze4 = Left(z4, 1)
z4 = z4 + (ze4 * 6)
End If
MSComm1.Output = Chr$(z4)
MSComm1.Output = Chr$(z3)
MSComm1.Output = Chr$(z2)
MSComm1.Output = Chr$(z1)
MSComm1.Output = Chr$(10)
 
For the meter incoming signal and power transmit output data reading the app use a timer set at 500 ms to read periodical the actual transceiver status information. The program need to analyze the dump of data and extract the proper values. A progress bargraph control display the incoming value. The meter status data consist of four identical bytes followed by a filler byte. The incoming metering value range goes from 1 to 255 (8 bits). The On Error procedure avoid the application hang up in case of accidental RS232 serial cable disconnection.
Private Sub Timer2_Timer()
On Error GoTo error
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(0)
MSComm1.Output = Chr$(247)
For a = 1 To 5000000
Next a
For a = 1 To 5
buffer$ = buffer$ & MSComm1.Input
Next a
b$ = Left(buffer$, 1)
c = Asc(b$)
Barra.Value = c
Exit Sub
error:
Barra.Value = 1
End Sub
 
Every time if a command button on the dashboard panel is pressed a beep sound is played to inform to the user the event. For this procedure the MCI multimedia control is used. The next code select the tone.wav file to reproduce a 1khz beep sound. 
Private Sub Form_Load()
Timer2.Enabled = False
MCI.DeviceType = "WaveAudio"
MCI.FileName = App.Path + "\Tone.wav"
MCI.Wait = True
MCI.Notify = False
MCI.Command = "Open"
MCI.UpdateInterval = 100
End Sub
 
In this brief tutorial the most important sections of the code in Visual Basic 6 are covered. Of course the program can be improved adding for example a menu plus a small database for the storage of memories and CQ contacts, etc. The core functions for I/O remote operation are completed and operational. Also the application is working without critical issues, an important aspect in software development. Of course adding more functions and features is my next task. How to resolve the data and parameters conversions on the project under Visual Basic 6 using Yaesu CAT protocol interface was the objective.

