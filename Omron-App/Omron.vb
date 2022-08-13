
#Region "Options"

Option Explicit On

#End Region

#Region "References"

Imports System.IO.Ports

#End Region

Namespace Omron

#Region "Omrom Classes"

#Region "Omron Exception"

    '''<summary>
    '''Omron Exception Class
    ''' Version 2.0.1
    ''' Created and Tested by: Mahdi Mansouri
    ''' Farzankar Ind. Co.
    ''' Apr 04 2011
    ''' </summary>
    Friend Class OmronException
        Inherits Exception
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub

    End Class

#End Region

#Region "Omron PLC"

    ''' <summary>
    ''' Omron PLC Interface Class
    ''' Version 2.0.1
    ''' Created and Tested by: Mahdi Mansouri
    ''' Farzankar Ind. Co.
    ''' Apr 04 2011
    ''' </summary>
    Friend Class OmronPLC

#Region "Fields"

        Private byte_time As Single = -1
        Private Const delay_ticks As Long = TimeSpan.TicksPerMillisecond * 100
        Private Const delay_ms As Long = 100
        Private inuse_String As ArrayList = New ArrayList()
        Private is_open As Boolean = False
        Private is_alive As Boolean = False
        Private Link As SerialPort = Nothing
        Private p_filepath As String = App_Path() & "omron.plc"
        Friend Const Size As Integer = 4 * 5 + 2 * 4
        Private Shared str_Command_Header As String() = New String() {"TS", "MS", "MF", _
        "RR", "RH", "RJ", "RL", "RG", "RD", "RX", "RF", "RC", "R#", "R$", "R%", _
        "SC", "WR", "WH", "WJ", "WL", "WG", "WD", "WF", "WC", "W#", "W$", "W%", _
        "KS", "KR", "FK", "FR", "KC", _
        "MM", _
        "CR", _
        "XZ", _
        "EX", _
        "IC", _
        "  ", _
        "RP", "RI", "WP", "MI", "QQ", "QQ"}
        Private Shared str_Command_SubHeader As String() = New String() {"MR", "IR"}
        Private unit_Address As Integer = 0

#End Region

#Region "Methods"

        ''' <summary>
        ''' Contains Last \ sign, e.g. App_Path and test.txt
        ''' </summary>
        Private Function App_Path() As String
            Return System.AppDomain.CurrentDomain.BaseDirectory()
        End Function

        ''' <summary>
        ''' Add an Reading String Request to Buffer Array
        ''' </summary>
        Public Function BArr_ReadAdd(ByVal memory_area As Omron_Command_Header_Read, ByVal start_address As Integer, ByVal count As Integer) As Boolean
            Return (BArr_ReadAdd(UnitAddress, memory_area, start_address, count))
        End Function

        ''' <summary>
        ''' Add an Reading String Request to Buffer Array
        ''' </summary>
        Public Function BArr_ReadAdd(ByVal unit_add As Integer, ByVal memory_area As Omron_Command_Header_Read, ByVal start_address As Integer, ByVal count As Integer) As Boolean
            If (start_address + count) > 9999 Or start_address < 0 Then Throw New OmronException("Start Adrress is not valid!")
            Dim str_out As String = "@" & _
                        Math.Abs(unit_add Mod 32).ToString.PadLeft(2, "0"c) & _
                        str_Command_Header(memory_area) & _
                        (start_address).ToString.PadLeft(4, "0"c) & _
                        count.ToString.PadLeft(4, "0"c)
            str_out = str_out & FCSCode_Set(str_out) & "*" & vbCr
            inuse_String.Add(str_out)
            Return True
        End Function

        ''' <summary>
        ''' Remove an element from the Buffer Array
        ''' </summary>
        Public Function BArr_RemoveAt(ByVal index As String) As Boolean
            Try
                If index >= BArrCount - 1 Then Throw New OmronException("Index is not in range!")
                inuse_String.RemoveAt(index)
                Return False
            Catch ex As Exception
                Throw New OmronException(ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Conver Bool to Byte
        ''' </summary>
        Protected Shared Function Bool2byte(ByVal arr_bool As Boolean) As Byte
            If arr_bool Then Return 1
            Return 0
        End Function

        ''' <summary>
        ''' Conver Bool Array to Byte
        ''' </summary>
        Protected Shared Function Bool2byte(ByVal arr_bool As Boolean()) As Byte()
            Dim arr_out As Byte() = New Byte() {}
            If arr_bool Is Nothing Then Return arr_out
            arr_out = New Byte(UBound(arr_bool) \ 8) {}
            Dim I As Integer = LBound(arr_bool), u_b As Integer = UBound(arr_bool)
            Do
                arr_out(I \ 8) = arr_out(I \ 8) Or IIf(arr_bool(I), 2 ^ (I Mod 8), 0)
                I += 1
            Loop While (Not (I > u_b))
            Return arr_out
        End Function

        ''' <summary>
        ''' Conver Bool to Word
        ''' </summary>
        Protected Shared Function Bool2Word(ByVal arr_bool As Boolean) As UShort
            If arr_bool Then Return 1
            Return 0
        End Function

        ''' <summary>
        ''' Conver Bool Array to Word
        ''' </summary>
        Protected Shared Function Bool2Word(ByVal arr_bool As Boolean()) As UShort()
            Dim arr_out As UShort() = New UShort() {}
            If arr_bool Is Nothing Then Return arr_out
            arr_out = New UShort(UBound(arr_bool) \ 16) {}
            Dim I As Integer = LBound(arr_bool), u_b As Integer = UBound(arr_bool)
            Do
                arr_out(I \ 16) = arr_out(I \ 16) Or IIf(arr_bool(I), 2 ^ (I Mod 16), 0)
                I += 1
            Loop While (Not (I > u_b))
            Return arr_out
        End Function

        ''' <summary>
        ''' Conver BoolArr to Hex
        ''' </summary>
        Protected Shared Function BoolArr2Hex(ByVal arr_bool As Boolean()) As String
            Dim str_out As String = ""
            Dim arr_size As Integer = IIf((arr_bool.Length Mod 16) = 0, arr_bool.Length - 1, ((1 + (arr_bool.Length \ 16)) * 16) - 1)
            Dim arr_bool_temp As Boolean() = New Boolean(arr_size) {}
            Array.Copy(arr_bool, LBound(arr_bool), arr_bool_temp, 0, arr_bool.Length)
            Dim I As Integer = 0, u_b As Integer = UBound(arr_bool_temp)
            Dim arr_temp As Boolean() = New Boolean(15) {}
            Do
                Array.Copy(arr_bool_temp, I, arr_temp, 0, 16)
                str_out = str_out & UShort2hex(Bool2Word(arr_temp)(0))
                I += 16
            Loop While (Not (I > u_b))
            Return str_out
        End Function

        ''' <summary>
        ''' Conver Byte to Hex
        ''' </summary>
        Protected Shared Function Byte2Hex(ByVal value As Byte) As String
            Return Hex(value).PadLeft(2, "0"c)
        End Function

        ''' <summary>
        ''' Conver ByteArr to Hex
        ''' </summary>
        Protected Shared Function ByteArr2Hex(ByVal arr_value As Byte()) As String
            Dim I As Integer = 0
            Dim arr_size As Integer = IIf((arr_value.Length Mod 2) = 0, arr_value.Length - 1, arr_value.Length)
            Dim arr_byte As Byte() = New Byte(arr_size) {}
            Array.Copy(arr_value, LBound(arr_value), arr_byte, 0, arr_value.Length)
            Dim str_out As String = ""
            Dim u_b As Integer = UBound(arr_byte)
            Do
                str_out = str_out & Byte2Hex(arr_byte(I))
                I += 1
            Loop While (Not (I > u_b))
            Return str_out
        End Function

        ''' <summary>
        ''' Conver Char to Byte
        ''' </summary>
        Protected Shared Function Char2Byte(ByVal element As Char) As Byte
            Return CByte(AscW(element))
        End Function

        ''' <summary>
        ''' Close CommPort
        ''' </summary>
        Public Function Close() As Boolean
            If Link Is Nothing Then Throw New OmronException("Object is not initialized yet!")
            Try
                If Link.IsOpen Then Link.Close()
            Catch ex As Exception
                Throw New OmronException("CommPort Error Occurred!" & vbNewLine & ex.Message)
            End Try
            Alive = False
            Opened = False
        End Function

        ''' <summary>
        ''' Copy from another object
        ''' </summary>
        Public Function CopyFrom(ByVal plc As Omron.OmronPLC) As Boolean
            If plc Is Nothing Then Throw New OmronException("The source object is null!")
            If plc.CommSetting.IsOpen Then Throw New OmronException("Copying while the source object port is open prohibited!")
            If CommSetting.IsOpen Then Throw New OmronException("Copying while the destination object port is open prohibited!")
            With plc

                Me.CommSetting.BaudRate = .CommSetting.BaudRate
                Me.CommSetting.DataBits = .CommSetting.DataBits
                Me.CommSetting.Parity = .CommSetting.Parity
                Me.CommSetting.PortName = .CommSetting.PortName
                Me.CommSetting.StopBits = .CommSetting.StopBits

                Me.UnitAddress = .UnitAddress
            End With
            Return True
        End Function

        ''' <summary>
        ''' Convert Double to Hex
        ''' </summary>
        Protected Shared Function Double2Hex(ByVal Value As Double) As String
            Try
                Dim ByteArr(7) As Byte
                ByteArr = BitConverter.GetBytes(Value)
                Double2Hex = ByteArr(1).ToString("X2") & ByteArr(0).ToString("X2") & ByteArr(3).ToString("X2") & ByteArr(2).ToString("X2") & ByteArr(5).ToString("X2") & ByteArr(4).ToString("X2") & ByteArr(7).ToString("X2") & ByteArr(6).ToString("X2")
            Catch ex As Exception
                Return ("0000000000000000")
            End Try
        End Function

        ''' <summary>
        ''' Convert DoubleArr to Hex
        ''' </summary>
        Protected Shared Function DoubleArr2Hex(ByVal arr_double As Double()) As String
            Dim I As Integer = LBound(arr_double), u_b As Integer = UBound(arr_double)
            Dim str_out As String = ""
            Do
                str_out = str_out & Double2Hex(arr_double(I))
                I += 1
            Loop While Not (I > u_b)
            Return str_out
        End Function

        ''' <summary>
        ''' Generates String Decode Errors
        ''' </summary>
        Public Shared Function [Error](ByVal p_response As Omron.Omron_Response_Code, ByVal uByte As Omron.Omron_Error_UpperByte, ByVal mByte As Omron.Omron_Error_MiddleByte, ByVal lByte As Omron.Omron_Error_LowerByte) As String()
            Dim str_err As String() = New String(19) {}
            str_err(0) = Response(p_response)(0)

            'uByte Error Response
            str_err(1) = ("I/O Buss Error (C0 to 4) is " & IIf(uByte.Err_C0 = Omron.Omron_Error_C04_to_04_IO_Bus_Error.off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(2) = ("End Instruction Missing (F0) is " & IIf(uByte.Err_F0 = Omron.Omron_Error_F0_End_Instruction_Missing.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(3) = ("Memory Error (F1) is " & IIf(uByte.Err_F1 = Omron.Omron_Error_F1_Memory_Error.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(4) = ("Jump Instruction Error (F2) is " & IIf(uByte.Err_F2 = Omron.Omron_Error_F2_Jump_Instruction_Error.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(5) = ("Program Error (F3) is " & IIf(uByte.Err_F3 = Omron.Omron_Error_F3_Program_Error.off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(6) = ("RTI Instruction Error (F4) is " & IIf(uByte.Err_F4 = Omron.Omron_Error_F4_RTI_Instruction_Error.[Off], "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(7) = ("FALS(CPU Stops) Error is " & IIf(uByte.Err_FALS = Omron.Omron_Error_FALS.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(8) = ("Host Link Unit Transmission Error is " & IIf(uByte.Err_HostLink = Omron.Omron_Error_Host_Link_Unit_Transmission_Error.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(9) = ("PC Link Error is " & IIf(uByte.Err_PCLink = Omron.Omron_Error_PC_Link_Error.Off, "Off", "Actived")).ToString.PadRight(80, " ")

            'mByte Error Response
            Select Case mByte.Err_Data_IOBus
                Case Omron.Omron_Error_Data_From_IO_Bus.Group_1_Control_Signal_Error
                    str_err(10) = ("Data From IO Bus is Address Bus Failure(Group1)").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Data_From_IO_Bus.Group_2_Data_Bus_Failure
                    str_err(10) = ("Data From IO Bus is Address Bus Failure(Group2)").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Data_From_IO_Bus.Group_3_Address_Bus_Failure
                    str_err(10) = ("Data From IO Bus is Address Bus Failure(Group3)").ToString.PadRight(80, " ")
                Case Else
                    str_err(10) = ("No Error Data From IO Bus").ToString.PadRight(80, " ")
            End Select
            str_err(11) = ("Duplex Bus Error C2000H Only Intelligent IO Error C200H HS only " & IIf(mByte.Err_Dup_IO = Omron.Omron_Error_Duplex_Bus_Error_C2000H_Only_Intelligent_IO_Error_C200H_HS_only.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(12) = ("Battery Failure (F7) is " & IIf(mByte.Err_F7_Battery = Omron.Omron_Error_F7_Battery_Failure.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(13) = ("FAL Error is " & IIf(mByte.Err_FAL = Omron.Omron_Error_FAL.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            Select Case mByte.Err_UnitNo
                Case Omron.Omron_Error_Unit_Number.Rack_01
                    str_err(14) = ("Error Occurred in Rack#01").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Unit_Number.Rack_02
                    str_err(14) = ("Error Occurred in Rack#02").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Unit_Number.Rack_03
                    str_err(14) = ("Error Occurred in Rack#03").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Unit_Number.Rack_04
                    str_err(14) = ("Error Occurred in Rack#04").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Unit_Number.Rack_05
                    str_err(14) = ("Error Occurred in Rack#05").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Unit_Number.Rack_06
                    str_err(14) = ("Error Occurred in Rack#06").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Unit_Number.Rack_07
                    str_err(14) = ("Error Occurred in Rack#07").ToString.PadRight(80, " ")
                Case Omron.Omron_Error_Unit_Number.Rack_CPU
                    str_err(14) = ("Error Occurred in CPU Rack").ToString.PadRight(80, " ")
            End Select

            'lByte Error Response
            str_err(15) = ("Remote IO Error (B0 to 03) is " & IIf(lByte.Err_B0 = Omron.Omron_Error_B0_To_B3_Remote_IO_Error.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(16) = ("IO Setting Error (E0) is " & IIf(lByte.Err_E0 = Omron.Omron_Error_E0_IO_Setting_Error.off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(17) = ("IO Unit Over (E1) is " & IIf(lByte.Err_E1 = Omron.Omron_Error_E1_IO_Unit_Over.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(18) = ("IO Verify Error (F7) is " & IIf(lByte.Err_F7 = Omron.Omron_Error_F7_IO_Verify_Error.Off, "Off", "Actived")).ToString.PadRight(80, " ")
            str_err(19) = ("Cycle Timer Over (F8) is " & IIf(lByte.Err_F8 = Omron.Omron_Error_F8_Cycle_Timer_Over.Off, "Off", "Actived")).ToString.PadRight(80, " ")

            Return str_err
        End Function

        ''' <summary>
        ''' Verify FCS Code validity
        ''' </summary>
        Protected Shared Function FCSCode_Get(ByVal Message As String, Optional ByVal is_last As Boolean = False) As Boolean
            'Private Function FCSCode_Get(ByVal Message As String) As Boolean
            'Works whith Complete or Incomplete messages includig/not including *
            Dim fcr As String
            Dim str_temp As String
            Message = Message.Trim()
            If Message.Length < 10 Then
                Return False
            End If
            If Message.Substring(Message.Length - 1, 1) = "*" Then
                fcr = Message.Substring(Message.Length - 3, 2)
                str_temp = Message.Substring(0, Message.Length - 3)
            Else
                fcr = Message.Substring(Message.Length - 2, 2)
                str_temp = Message.Substring(0, Message.Length - 2)
            End If
            Return (fcr = FCSCode_Set(str_temp, is_last))
        End Function

        ''' <summary>
        ''' Create FCS for current String
        ''' </summary>
        Protected Shared Function FCSCode_Set(ByVal Message As String, Optional ByVal is_last As Boolean = False) As String
            'Works with uncompleted message , Excluding (@,FCS,*,CR)
            Dim [xor] As Integer = IIf(is_last, 0, 64) '@ascii code
            For I As Integer = 1 To Message.Length - 1
                [xor] = [xor] Xor Convert.ToInt32(Convert.ToChar(Message.Substring(I, 1)))
            Next
            Return ([String].Format("{0:X}", [xor]))
        End Function

        ''' <summary>
        ''' Dispose Current Object
        ''' </summary>
        Protected Overrides Sub Finalize()
            If ISOpen Then Me.Close()
            Alive = False
            Opened = False
            Link = Nothing
            If Not inuse_String Is Nothing Then inuse_String.Clear()
            inuse_String = Nothing
            MyBase.Finalize()
        End Sub

        ''' <summary>
        ''' Convert Hexadecimal to Byte
        ''' </summary>
        Protected Shared Function Hex2Byte(ByVal Data As String) As Byte
            Dim value As Integer
            value = Convert.ToByte(Data.Substring(0, 2), 16)
            Return CByte(value)
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Bytee Array
        ''' </summary>
        Protected Shared Function Hex2ByteArr(ByVal str_data As String) As Byte()
            Dim I As Integer = 1
            Dim arr_size As Integer = (str_data.Length \ 2) * 2 + IIf((str_data.Length Mod 2) = 0, 0, 2)
            str_data = str_data.PadLeft(arr_size, "0"c)
            arr_size = arr_size / 2
            Dim arr_byte As Byte() = New Byte(arr_size - 1) {}
            While Not (I > str_data.Length)
                arr_byte((I - 1) / 2) = Hex2Byte(Mid$(str_data, I, 2))
                I += 2
            End While
            Return arr_byte
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Bool Array
        ''' </summary>
        Protected Shared Function Hex2BoolArr(ByVal str_data As String) As Boolean()
            If Trim(str_data) = "" Then Return New Boolean() {False}
            Dim J As Integer = 0, I As Integer = 1, Offset As Integer = 0
            Dim temp_short As UShort = 0
            Dim arr_size As Integer = (str_data.Length \ 4) * 4 + IIf((str_data.Length Mod 4) = 0, 0, 4)
            str_data = str_data.PadLeft(arr_size, "0"c)
            arr_size = (arr_size * 4) - 1
            Dim arr_bool As Boolean() = New Boolean(arr_size) {}
            While Not (I > str_data.Length)
                temp_short = Hex2Word(Mid$(str_data, I, 4))
                J = 0
                Offset = 4 * (I - 1)
                While J < 16
                    arr_bool(Offset + J) = (temp_short And (2 ^ J)) >> J
                    J += 1
                End While
                I += 4
            End While
            Return arr_bool
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Double
        ''' </summary>
        Protected Shared Function Hex2Double(ByVal hexValue As String) As Double
            Try
                hexValue = hexValue.Substring(0, 16)
                Dim iInputIndex As Integer = 0
                Dim iOutputIndex As Integer = 0
                Dim bArray(7) As Byte
                For iInputIndex = hexValue.Length - 2 To 0 Step -2
                    bArray(iOutputIndex) = Byte.Parse(hexValue.Chars(iInputIndex) & hexValue.Chars(iInputIndex + 1), Globalization.NumberStyles.HexNumber)
                    iOutputIndex += 1
                Next
                Array.Reverse(bArray)
                Return BitConverter.ToDouble(bArray, 0)
            Catch ex As Exception
                Return (0.0)
            End Try
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Double Array
        ''' </summary>
        Protected Shared Function Hex2DoubleArr(ByVal str_data As String) As Double()
            Dim I As Integer = 1
            Dim arr_size As Integer = (str_data.Length \ 16) * 16 + IIf((str_data.Length Mod 16) = 0, 0, 16)
            str_data = str_data.PadLeft(arr_size, "0"c)
            arr_size = arr_size / 16
            Dim arr_double As Double() = New Double(arr_size - 1) {}
            While Not (I > str_data.Length)
                arr_double(I / 16) = Hex2Double(Mid$(str_data, I, 16))
                I += 16
            End While
            Return arr_double
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Integer
        ''' </summary>
        Protected Shared Function Hex2Int(ByVal hexvalue As String) As Integer
            Try
                hexvalue = hexvalue.Substring(0, 8)
                Dim iInputIndex As Integer = 0
                Dim iOutputIndex As Integer = 0
                Dim bArray(3) As Byte
                For iInputIndex = hexValue.Length - 2 To 0 Step -2
                    bArray(iOutputIndex) = Byte.Parse(hexValue.Chars(iInputIndex) & hexValue.Chars(iInputIndex + 1), Globalization.NumberStyles.HexNumber)
                    iOutputIndex += 1
                Next
                Array.Reverse(bArray)
                Return BitConverter.ToInt32(bArray, 0)
            Catch ex As Exception
                Return (0)
            End Try
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Integer Array
        ''' </summary>
        Protected Shared Function Hex2IntArr(ByVal str_data As String) As Integer()
            Dim I As Integer = 1
            Dim data As Integer = (str_data.Length \ 8) * 8 + IIf((str_data.Length Mod 8) = 0, 0, 8)
            str_data = str_data.PadLeft(data, "0"c)
            data = data / 8
            Dim arr_integer As Integer() = New Integer(data - 1) {}
            While Not (I > str_data.Length)
                arr_integer(I / 8) = Hex2Int(Mid$(str_data, I, 8))
                I += 8
            End While
            Return arr_integer
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Short
        ''' </summary>
        Protected Shared Function Hex2Short(ByVal Data As String) As Short
            Dim value As Integer
            value = Convert.ToInt16(Data, 16)
            Return CShort(value)
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Short Array
        ''' </summary>
        Protected Shared Function Hex2ShortArr(ByVal str_Data As String) As Short()
            Dim I As Integer = 1
            Dim data As Integer = (str_Data.Length \ 4) * 4 + IIf((str_Data.Length Mod 4) = 0, 0, 4)
            str_Data = str_Data.PadLeft(data, "0"c)
            data = data / 4
            Dim arr_short As Short() = New Short(data - 1) {}
            While Not (I > str_Data.Length)
                arr_short(I / 4) = Hex2Short(Mid$(str_Data, I, 4))
                I += 4
            End While
            Return arr_short
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Single
        ''' </summary>
        Protected Shared Function Hex2Single(ByVal hexValue As String) As Single
            Try
                Dim iInputIndex As Integer = 0
                Dim iOutputIndex As Integer = 0
                Dim bArray(3) As Byte
                For iInputIndex = hexValue.Length - 2 To 0 Step -2
                    bArray(iOutputIndex) = Byte.Parse(hexValue.Chars(iInputIndex) & hexValue.Chars(iInputIndex + 1), Globalization.NumberStyles.HexNumber)
                    iOutputIndex += 1
                Next
                Array.Reverse(bArray)
                Return BitConverter.ToSingle(bArray, 0)
            Catch ex As Exception
                Return (0.0)
            End Try
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Single Array
        ''' </summary>
        Protected Shared Function Hex2SingleArr(ByVal str_data As String) As Single()
            Dim I As Integer = 1
            Dim data As Integer = (str_data.Length \ 8) * 8 + IIf((str_data.Length Mod 8) = 0, 0, 8)
            str_data = str_data.PadLeft(data, "0"c)
            data = data / 8
            Dim arr_single As Single() = New Single(data - 1) {}
            While Not (I > str_data.Length)
                arr_single(I / 8) = Hex2Single(Mid$(str_data, I, 8))
                I += 8
            End While
            Return arr_single
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Word
        ''' </summary>
        Protected Shared Function Hex2Word(ByVal Data As String) As UShort
            Dim value As Integer
            value = Convert.ToUInt16(Data, 16)
            Return CUShort(value)
        End Function

        ''' <summary>
        ''' Convert Hexadecimal to Word
        ''' </summary>
        Protected Shared Function Hex2WordArr(ByVal str_Data As String) As UShort()
            Dim I As Integer = 1
            Dim data As Integer = (str_Data.Length \ 4) * 4 + IIf((str_Data.Length Mod 4) = 0, 0, 4)
            str_Data = str_Data.PadLeft(data, "0"c)
            data = data / 4
            Dim arr_ushort As UShort() = New UShort(data - 1) {}
            While Not (I > str_Data.Length)
                arr_ushort(I / 4) = Hex2Word(Mid$(str_Data, I, 4))
                I += 4
            End While
            Return arr_ushort
        End Function

        ''' <summary>
        ''' Convert Integer to Hexadecimal
        ''' </summary>
        Protected Shared Function Int2Hex(ByVal Data As Integer, Optional ByVal PadNum As Integer = 8) As String
            Try
                Dim ByteArr(4) As Byte
                ByteArr = BitConverter.GetBytes(Data)
                Int2Hex = ByteArr(1).ToString("X2") & ByteArr(0).ToString("X2") & ByteArr(3).ToString("X2") & ByteArr(2).ToString("X2")
                Int2Hex = Int2Hex.PadLeft(PadNum, "0"c)
            Catch ex As Exception
                Return ("00000000")
            End Try
        End Function

        ''' <summary>
        ''' Convert Integer to Hexadecimal
        ''' </summary>
        Protected Shared Function IntArr2Hex(ByVal arr_int As Integer()) As String
            Dim I As Integer = LBound(arr_int), u_b As Integer = UBound(arr_int)
            Dim str_out As String = ""
            Do
                str_out = str_out & Int2Hex(arr_int(I))
                I += 1
            Loop While Not (I > u_b)
            Return str_out
        End Function

        ''' <summary>
        ''' Convert Int32 to Hex
        ''' </summary>
        Protected Shared Function Integer2Hex(ByVal IntegerValue As Integer) As String
            Try
                Dim ByteArr(4) As Byte
                ByteArr = BitConverter.GetBytes(IntegerValue)
                Integer2Hex = ByteArr(1).ToString("X2") & ByteArr(0).ToString("X2") & ByteArr(3).ToString("X2") & ByteArr(2).ToString("X2")
            Catch ex As Exception
                Return ("00000000")
            End Try
        End Function

        ''' <summary>
        ''' Convert Int32Arr to Hex
        ''' </summary>
        Protected Shared Function IntegerArr2Hex(ByVal arr_int As Integer()) As String
            Dim I As Integer = LBound(arr_int), u_b As Integer = UBound(arr_int)
            Dim str_out As String = ""
            Do
                str_out = str_out & Integer2Hex(arr_int(I))
                I += 1
            Loop While Not (I > u_b)
            Return str_out
        End Function

        ''' <summary>
        ''' Check Whether Current String a frame is or not
        ''' </summary>
        Protected Shared Function ISFrame(ByVal Message As String) As Boolean
            'Works whith whole message includig FCR, *, CR
            Dim result As Boolean = True
            If Message.Substring(0, 1) <> "@" Then
                result = result And False
            ElseIf Message.Substring(Message.Trim().Length - 1, 1) <> "*" Then
                result = result And False
            End If
            Return result
        End Function

        ''' <summary>
        ''' Loads Object vaues from a file on Hard Disk
        ''' </summary>
        Public Function Load() As Boolean
            Return Load(Path)
        End Function

        ''' <summary>
        ''' Loads Object values from a file on Hard Disk
        ''' </summary>
        Public Function Load(ByVal f_path As String) As Boolean
            If ISOpen Then Throw New OmronException("Please close the port before loading file into!")
            If Link Is Nothing Then Link = New IO.Ports.SerialPort()

            Dim FileInput As IO.FileStream = Nothing
            Dim BinaryReader As IO.BinaryReader = Nothing
            Dim I As Integer = 0
            Try
                FileInput = New IO.FileStream(f_path, IO.FileMode.Open, IO.FileAccess.Read)
                BinaryReader = New IO.BinaryReader(FileInput)
                FileInput.Seek(0, 0)

                With CommSetting
                    .BaudRate = BinaryReader.ReadInt32
                    .DataBits = BinaryReader.ReadInt32
                    .Parity = BinaryReader.ReadInt32
                    .PortName = Trim$(BinaryReader.ReadString)
                    .StopBits = BinaryReader.ReadInt32
                End With
                UnitAddress = BinaryReader.ReadInt32

                If Not BinaryReader Is Nothing Then BinaryReader.Close()
                If Not FileInput Is Nothing Then FileInput.Close()
                Return True
            Catch ex As Exception
                If Not BinaryReader Is Nothing Then BinaryReader.Close()
                If Not FileInput Is Nothing Then FileInput.Close()
                Throw New OmronException(ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Create New Instance of Object
        ''' </summary>
        Sub New()
            Call Me.New(New SerialPort)
        End Sub

        ''' <summary>
        ''' Create New Instance of Object
        ''' </summary>
        Sub New(ByVal serial As SerialPort)
            Call Me.New(serial, 0)
        End Sub

        ''' <summary>
        ''' Create New Instance of Object
        ''' </summary>
        Sub New(ByVal serial As SerialPort, ByVal unit_address As Integer)
            If serial Is Nothing Then serial = New SerialPort
            CommSetting = serial
            UnitAddress = unit_address
        End Sub

        ''' <summary>
        ''' Open CommPort for Communication
        ''' </summary>
        Public Function Open() As Boolean
            If Link Is Nothing Then Throw New OmronException("Object is not initialized yet!")
            Try
                If Not Link.IsOpen Then Link.Open()
            Catch ex As Exception
                Throw New OmronException("CommPort Error Occurred!" & vbNewLine & ex.Message)
            End Try
            Opened = True
        End Function

        ''' <summary>
        ''' Read from PLC
        ''' </summary>
        Public Function Read(ByVal buffer_index As Integer, ByRef p_response As Omron.Omron_Response_Code, Optional ByVal return_type As Omron_Data_Type = Omron_Data_Type.Char) As System.Array
            Dim count As Integer = 0
            count = CInt(Val(Mid$(BArr(buffer_index), 10, 4)))
            Return (Read(BArr(buffer_index), count, p_response, return_type))
        End Function

        ''' <summary>
        ''' Read from PLC
        ''' </summary>
        Public Function Read(ByVal memory_area As Omron_Command_Header_Read, ByVal start_address As Integer, ByVal count As Integer, ByRef p_response As Omron.Omron_Response_Code, Optional ByVal return_type As Omron_Data_Type = Omron_Data_Type.Char) As System.Array
            Return (Read(UnitAddress, memory_area, start_address, count, p_response, return_type))
        End Function

        ''' <summary>
        ''' Read from PLC
        ''' </summary>
        Public Function Read(ByVal unit_add As Integer, ByVal memory_area As Omron_Command_Header_Read, ByVal start_add As Integer, ByVal count As Integer, ByRef p_response As Omron.Omron_Response_Code, Optional ByVal return_type As Omron_Data_Type = Omron_Data_Type.Char) As System.Array

            If (start_add + count) > 9999 Or start_add < 0 Then Throw New OmronException("Start Adrress is not valid!")
            Dim str_out As String = "@" & _
                Math.Abs(unit_add Mod 32).ToString.PadLeft(2, "0"c) & _
                str_Command_Header(memory_area) & _
                (start_add).ToString.PadLeft(4, "0"c) & _
                count.ToString.PadLeft(4, "0"c)
            str_out = str_out & FCSCode_Set(str_out) & "*" & vbCr
            Return (Read(str_out, count, p_response, return_type))

        End Function

        ''' <summary>
        ''' Read from PLC
        ''' </summary>
        Protected Function Read(ByVal str_out As String, ByVal count As Integer, ByRef p_response As Omron.Omron_Response_Code, Optional ByVal return_type As Omron_Data_Type = Omron_Data_Type.Char) As System.Array

            If Link Is Nothing Then Throw New OmronException("The object is not initilized yet!")
            If Not ISOpen Then Throw New OmronException("Port is not opened")
            If count > 9999 Or count < 1 Then Throw New OmronException("Number to be read is not in range!")

            Dim loop_count As Integer = 0, count_predicted As Integer = 0, _
            lng_length As Long = 0, tmr_indicator As Long = 0, _
            length_predicted As Integer = 0
            Dim str_in As String = "", str_in_final As String = ""

            str_in = Link.ReadExisting()
            str_in = ""
            Link.Write(str_out)

            Do
                str_in = ""
                count_predicted = IIf(count > 30, 30, count)
                If loop_count = 0 Then
                    If count_predicted = count Then
                        length_predicted = count_predicted * 4 + 11
                    Else
                        length_predicted = count_predicted * 4 + 10
                    End If
                Else
                    If count_predicted = count Then
                        length_predicted = count_predicted * 4 + 4
                    Else
                        length_predicted = count_predicted * 4 + 3
                    End If
                End If

                lng_length = (length_predicted * ByteTime * 1000) + delay_ms
                lng_length *= TimeSpan.TicksPerMillisecond

                tmr_indicator = DateAndTime.Now.Ticks

                Do
                    str_in = str_in & Link.ReadExisting
                Loop While ((str_in.Length < length_predicted) And ((DateAndTime.Now.Ticks - tmr_indicator) < lng_length))

                If ((loop_count > 0) And (str_in.Length < length_predicted)) Then Throw New OmronException("Error in Data line, Data corrupted!")

                If loop_count = 0 Then 'Indicates first frame
                    If Mid$(str_in, str_in.Length, 1) <> vbCr Then Throw New OmronException("Error in Data line, Data corrupted!")
                    If Not (FCSCode_Get(str_in)) Then Throw New OmronException("Comminucation Error, Please check the line!")
                    p_response = Hex2Int(Mid$(str_in, 6, 2))
                    If p_response <> Omron_Response_Code.Command_Completed_Normally Then Throw New OmronException(Response(p_response)(0))

                    If count_predicted = count Then 'One block contains the whole response
                        str_in_final = str_in_final & Mid$(str_in, 8, str_in.Length - 11)
                    Else 'It's the first block of a sequence of response blocks
                        str_in_final = str_in_final & Mid$(str_in, 8, str_in.Length - 10)
                    End If
                Else
                    If count_predicted = count Then 'Indicates last frame
                        str_in_final = str_in_final & Mid$(str_in, 1, str_in.Length - 4)
                    Else 'Indicates intermediate frame
                        str_in_final = str_in_final & Mid$(str_in, 1, str_in.Length - 3)
                    End If
                End If

                str_in = Link.ReadExisting
                str_in = ""
                If count_predicted < count Then Link.Write(vbCr)

                count = IIf(count - 30 > 0, count - 30, 0)
                loop_count += 1

            Loop While count > 0
            Dim arr_out As Array = Nothing
            If return_type = Omron_Data_Type.Boolean Then Return (Hex2BoolArr(str_in_final))
            If return_type = Omron_Data_Type.Byte Then Return (Hex2ByteArr(str_in_final))
            If return_type = Omron_Data_Type.Char Then Return (str_in_final.ToCharArray)
            If return_type = Omron_Data_Type.Double Then Return (Hex2DoubleArr(str_in_final))
            If return_type = Omron_Data_Type.Integer Then Return (Hex2IntArr(str_in_final))
            If return_type = Omron_Data_Type.Short Then Return (Hex2ShortArr(str_in_final))
            If return_type = Omron_Data_Type.Single Then Return (Hex2SingleArr(str_in_final))
            If return_type = Omron_Data_Type.UShort Then Return (Hex2WordArr(str_in_final))
            Return arr_out
        End Function

        ''' <summary>
        ''' Get PLC Error, Then Clear Errors
        ''' </summary>
        Public Function ReadError(ByRef p_response As Omron_Response_Code, ByRef U_byte As Omron_Error_UpperByte, ByRef M_byte As Omron_Error_MiddleByte, ByRef L_byte As Omron_Error_LowerByte, Optional ByVal Clear_Error As Boolean = False) As String
            Return (ReadError(UnitAddress, p_response, U_byte, M_byte, L_byte, Clear_Error))
        End Function

        ''' <summary>
        ''' Get PLC Error, Then Clear Errors
        ''' </summary>
        Public Function ReadError(ByVal unit_add As Integer, ByRef p_Response As Omron_Response_Code, ByRef U_byte As Omron_Error_UpperByte, ByRef M_byte As Omron_Error_MiddleByte, ByRef L_byte As Omron_Error_LowerByte, Optional ByVal Clear_Error As Boolean = True) As String
            If Link Is Nothing Then Throw New OmronException("The object is not initilized yet!")
            If Not ISOpen Then Throw New OmronException("Port is not opened")

            Dim lng_length As Integer, tmr_indicator As Long = 0, _
            length_predicted As Integer = 19
            Dim str_in As String, str_out As String = "@" & _
                Math.Abs(unit_add Mod 32).ToString.PadLeft(2, "0"c) & _
                "MF" & _
                IIf(Clear_Error, "01", "00")
            str_out = str_out & FCSCode_Set(str_out) & "*" & vbCr

            'Fixed Length for response of ReadError
            lng_length = (length_predicted * ByteTime * 1000) + delay_ms
            lng_length *= TimeSpan.TicksPerMillisecond

            str_in = Link.ReadExisting()
            str_in = ""

            Link.Write(str_out)

            tmr_indicator = DateAndTime.Now.Ticks
            Do
                str_in = str_in & Link.ReadExisting
            Loop While ((str_in.Length < length_predicted) And ((DateAndTime.Now.Ticks - tmr_indicator) < lng_length))

            If Not (FCSCode_Get(str_in)) Then Throw New OmronException("Comminucation Error, Please check the line!")

            'First, Extracting Responce Code
            p_Response = Hex2Int(str_in.Substring(6, 2))

            If p_Response <> Omron_Response_Code.Command_Completed_Normally Then Throw New OmronException(Response(p_Response)(0))
            'Second Step,Extracting First Word
            Dim err_code As Integer = Hex2Int(str_in.Substring(8, 4))

            U_byte.Err_FALS = (err_code And 2 ^ 15) >> 15
            U_byte.Err_F0 = (err_code And 2 ^ 14) >> 14
            U_byte.Err_F3 = U_byte.Err_F0
            U_byte.Err_HostLink = (err_code And 2 ^ 13) >> 13
            U_byte.Err_F4 = (err_code And 2 ^ 12) >> 12
            U_byte.Err_PCLink = (err_code And 2 ^ 11) >> 11
            U_byte.Err_C0 = (err_code And 2 ^ 10) >> 10
            U_byte.Err_F2 = (err_code And 2 ^ 9) >> 9
            U_byte.Err_F1 = (err_code And 2 ^ 8) >> 8

            M_byte.Err_FAL = (err_code And 2 ^ 7) >> 7
            M_byte.Err_Dup_IO = (err_code And 2 ^ 6) >> 6
            M_byte.Err_F7_Battery = (err_code And 2 ^ 5) >> 5
            M_byte.Err_UnitNo = (err_code And (2 ^ 4 Or 2 ^ 3 Or 2 ^ 2)) >> 2
            M_byte.Err_Data_IOBus = (err_code And (2 ^ 1 Or 2 ^ 0))

            err_code = Hex2Int(str_in.Substring(12, 2))

            L_byte.Err_F7 = (err_code And 2 ^ 3) >> 3
            L_byte.Err_F8 = (err_code And 2 ^ 4) >> 4
            L_byte.Err_E1 = (err_code And 2 ^ 5) >> 5
            L_byte.Err_E0 = (err_code And 2 ^ 6) >> 6
            L_byte.Err_B0 = (err_code And 2 ^ 7) >> 7

            Return ("FAL,FALS No. : " & Hex2Int(str_in.Substring(14, 2)).ToString())

        End Function

        ''' <summary>
        ''' Get PLC Status
        ''' </summary>
        Public Function ReadStatus(ByRef p_esponse As Omron_Response_Code, ByRef PLC_Status As Omron_Status) As String
            Return (ReadStatus(UnitAddress, p_esponse, PLC_Status))
        End Function

        ''' <summary>
        ''' Get PLC Status
        ''' </summary>
        Public Function ReadStatus(ByVal unit_add As Integer, ByRef p_response As Omron_Response_Code, ByRef PLC_Status As Omron_Status) As String
            If Link Is Nothing Then Throw New OmronException("The object is not initilized yet!")
            If Not ISOpen Then Throw New OmronException("Port is not opened")

            Dim lng_length As Integer, tmr_indicator As Long = 0, _
            length_predicted As Integer = 31
            Dim str_in As String, str_out As String = "@" & _
                Math.Abs(unit_add Mod 32).ToString.PadLeft(2, "0"c) & "MS"
            str_out = str_out & FCSCode_Set(str_out) & "*" & vbCr

            'Fixed Length for response of ReadStatus
            lng_length = (length_predicted * ByteTime * 1000) + delay_ms
            lng_length *= TimeSpan.TicksPerMillisecond

            str_in = Link.ReadExisting()
            str_in = ""
            Link.Write(str_out)

            tmr_indicator = DateAndTime.Now.Ticks
            Do
                str_in = str_in & Link.ReadExisting
            Loop While ((str_in.Length < length_predicted) And ((DateAndTime.Now.Ticks - tmr_indicator) < lng_length))

            If Not (FCSCode_Get(str_in)) Then Throw New OmronException("Comminucation Error, Please check the line!")

            'First, Extracting Responce Code
            p_response = Hex2Int(str_in.Substring(6, 2))
            If p_response <> Omron_Response_Code.Command_Completed_Normally Then Throw New OmronException(Response(p_response)(0))

            'Second Etracting PLC Status
            Dim int_responce As Integer = Hex2Int(str_in.Substring(8, 4))

            PLC_Status.False_Instruction = (int_responce And CInt(2 ^ 15)) >> 15
            PLC_Status.TransitionType = (int_responce And CInt(2 ^ 14)) >> 14
            PLC_Status.StartSwitch = (int_responce And CInt(2 ^ 13)) >> 13
            PLC_Status.Error_Diagnosis = (int_responce And CInt(2 ^ 12)) >> 12
            PLC_Status.Forced_Set_Reset = (int_responce And CInt(2 ^ 11)) >> 11
            PLC_Status.IO_Waiting = (int_responce And CInt(2 ^ 10)) >> 10
            PLC_Status.PLC_Mode = (int_responce And (CInt(2 ^ 9) Or CInt(2 ^ 8))) >> 8
            PLC_Status.Duplex_System = (int_responce And CInt(2 ^ 7)) >> 7
            PLC_Status.Program_Area_Size = (int_responce And (CInt(2 ^ 6) Or CInt(2 ^ 5) Or CInt(2 ^ 4))) >> 4
            PLC_Status.Program_Area = (int_responce And CInt(2 ^ 3)) >> 3

            If (str_in.Length < 16) Then
                Return "No Cmminucation Error"
            Else
                Return (str_in.Substring(12, str_in.Length - 15))
            End If
        End Function

        ''' <summary>
        ''' Generate Related Response String
        ''' </summary>
        Protected Shared Function Response(ByVal p_responce As Omron.Omron_Response_Code) As String()
            Dim str_resp As String() = Nothing
            str_resp = New String(0) {}

            Select Case p_responce
                Case Omron.Omron_Response_Code.Abort_Entry_Number_Err
                    str_resp(0) = "Aborted due to Entry number data error in transmit data".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Abort_FCS_Err
                    str_resp(0) = "Aborted due to FCS(Checksum) error in transmit data".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Abort_Format_Err
                    str_resp(0) = "Aborted due to Format error in transmit data".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Abort_Frame_Err
                    str_resp(0) = "Aborted due to Framing error in transmit data".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Abort_Frame_Length_Err
                    str_resp(0) = "Aborted due to Frame Length error in transmit data".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Abort_Overrun_Err
                    str_resp(0) = "Aborted due to Overrun error in transmit data".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Abort_Parity_Err
                    str_resp(0) = "Aborted due to Parity error in transmit data".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Address_Overflow
                    str_resp(0) = "Address Overflow(Data Overflow). Check the Program".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Command_Completed_Normally
                    str_resp(0) = "Command Completed Normally".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Command_Format_Err
                    str_resp(0) = "Command Format Error".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Exe_Blocked_Debug_Mode
                    str_resp(0) = "Execution was not possible because the PC is in DEBUG mode. Change the PC mode".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Exe_Blocked_HotLink_is_Local
                    str_resp(0) = "Execution was not possible because the HostLINK Unit's keyswitch is set to LOCAL mode or becuase the Command was sent to a C2000H CPU that was on Standby. Change the mode or send the command to the active CPU".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Exe_Blocked_Less_16KByte
                    str_resp(0) = "Not executable because the program area is not 16Kbytes".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Exe_Blocked_Monitor_Mode
                    str_resp(0) = "Execution was not possible because the PC is in Monitor mode. Change the PC mode".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Exe_Blocked_Program_Mode
                    str_resp(0) = "Execution was not possible because the PC is in Program mode. Change the PC mode".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Exe_Blocked_PROM_mounte
                    str_resp(0) = "Execution was not possible because PROM is mounted. Change the PC mode".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Exe_Blocked_Run_Mode
                    str_resp(0) = "Execution was not possible because the PC is in Run mode. Change the PC mode".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Exe_Blocked_UnExe_Err
                    str_resp(0) = "Execution was not possible because of an unexecutable error clear, memory error, EEPROM write disabled etc.".PadRight(80, " ")
                Case Omron.Omron_Response_Code.FCS_Err
                    str_resp(0) = "FCS error(Checksum Error)".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Frame_Err
                    str_resp(0) = "Framing error(Stop bit not detected)".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Frame_Length_Err
                    str_resp(0) = "Frame length error(Maximum length exceeded)".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Incorrect_Data_Area
                    str_resp(0) = "An incorrect Data Area designation was made for READ or WRITE".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Invalid_Instruction
                    str_resp(0) = "Instruction not found".PadRight(80, " ")
                Case Omron.Omron_Response_Code.IO_Table_Gen_Impossible
                    str_resp(0) = "I/O table generation was not possible(Unrecognized Remote I/O Unit, Word Overflow, Duplicated word allocation)".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Memory_not_Exist
                    str_resp(0) = "Memory Does not Exist".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Memory_Write_Protected
                    str_resp(0) = "Memory is Write Protected".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Overrun
                    str_resp(0) = "Overrun(the next command was received too soon)".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Parity_Err
                    str_resp(0) = "Parity Error".PadRight(80, " ")
                Case Omron.Omron_Response_Code.PC_s_CPU
                    str_resp(0) = "An error occurred in the PC's CPU".PadRight(80, " ")
                Case Omron.Omron_Response_Code.Unknown
                    str_resp(0) = "Unknown Error, Remover any possible causes of noise and resend the command".PadRight(80, " ")
                Case Else
                    str_resp(0) = "Unknown Error, Remover any possible causes of noise and resend the command".PadRight(80, " ")
            End Select
            Return str_resp
        End Function

        ''' <summary>
        ''' Saves Object into a file on Hard Disk
        ''' </summary>
        Friend Function Save() As Boolean
            Return Save(Path)
        End Function

        ''' <summary>
        ''' Saves Object into a file on Hard Disk
        ''' </summary>
        Friend Function Save(ByVal f_path As String) As Boolean
            Dim FileOutput As IO.FileStream = Nothing
            Dim BinaryOutput As IO.BinaryWriter = Nothing
            Dim I As Integer = 0
            Try
                FileOutput = New IO.FileStream(f_path, IO.FileMode.Create, IO.FileAccess.Write)
                FileOutput.SetLength(Size * 1)
                BinaryOutput = New IO.BinaryWriter(FileOutput)

                With CommSetting
                    BinaryOutput.Write(CInt(.BaudRate))
                    BinaryOutput.Write(CInt(.DataBits))
                    BinaryOutput.Write(CInt(.Parity))
                    BinaryOutput.Write(CStr(.PortName.PadRight(4, " "c).Substring(0, 4)))
                    BinaryOutput.Write(CInt(.StopBits))
                End With
                BinaryOutput.Write(CInt(UnitAddress))

                If Not FileOutput Is Nothing Then FileOutput.Close()
                If Not BinaryOutput Is Nothing Then BinaryOutput.Close()
                Return True
            Catch ex As Exception
                If Not FileOutput Is Nothing Then FileOutput.Close()
                If Not BinaryOutput Is Nothing Then BinaryOutput.Close()
                'Throw New OmronException(ex.Message)
            End Try

        End Function

        ''' <summary>
        ''' Convert Single to Double
        ''' </summary>
        Protected Shared Function Single2Double(ByVal value As Single) As Double
            Return CDbl(value)
        End Function

        ''' <summary>
        ''' Convert Single to Hex
        ''' </summary>
        Protected Shared Function Single2Hex(ByVal singleValue As Single) As String
            Try
                Dim ByteArr(3) As Byte
                ByteArr = BitConverter.GetBytes(singleValue)
                Single2Hex = ByteArr(1).ToString("X2") & ByteArr(0).ToString("X2") & ByteArr(3).ToString("X2") & ByteArr(2).ToString("X2")
            Catch ex As Exception
                Return ("00000000")
            End Try
        End Function

        ''' <summary>
        ''' Convert SingleArr to Hex
        ''' </summary>
        Protected Shared Function SingleArr2Hex(ByVal arr_single As Single()) As String
            Dim I As Integer = LBound(arr_single), u_b As Integer = UBound(arr_single)
            Dim str_out As String = ""
            Do
                str_out = str_out & Single2Hex(arr_single(I))
                I += 1
            Loop While Not (I > u_b)
            Return str_out
        End Function

        ''' <summary>
        ''' Convert Short to Hexadecimal
        ''' </summary>
        Protected Shared Function Short2Hex(ByVal Data As Short, Optional ByVal PadNum As Integer = 4) As String
            Dim hex As String = [String].Format("{0:X}", Data)
            Return hex.PadLeft(PadNum, "0"c)
        End Function

        ''' <summary>
        ''' Convert ShortArr to Hexadecimal
        ''' </summary>
        Protected Shared Function ShortArr2Hex(ByVal arr_short As Short()) As String
            Dim I As Integer = LBound(arr_short), u_b As Integer = UBound(arr_short)
            Dim str_out As String = ""
            Do
                str_out = str_out & Short2Hex(arr_short(I))
                I += 1
            Loop While Not (I > u_b)
            Return str_out
        End Function

        ''' <summary>
        ''' Generates String Decode Status
        ''' </summary>
        Public Shared Function Status(ByVal p_status As Omron.Omron_Status) As String()
            Dim str_status As String() = Nothing
            str_status = New String(4) {}


            str_status(0) = ("System Duplex System Status : " & IIf(p_status.Duplex_System = Omron.Omron_Duplex_System.Active, " Actived", "Standby")).PadRight(80, " ")
            str_status(1) = ("System Error Diagnosis Status : " & IIf(p_status.Error_Diagnosis = Omron.Omron_Error_Diagnosis.InProgress, "InProgress", "Off")).PadRight(80, " ")
            str_status(2) = ("System Instruction Status : " & IIf(p_status.False_Instruction, "False", "Correct")).PadRight(80, " ")
            str_status(3) = ("System Force Set/Reset Status : " & IIf(p_status.Forced_Set_Reset = Omron.Omron_Forced_Set_Reset.On, "On", "Off")).PadRight(80, " ")
            str_status(4) = ("System Remote IO Waiting Status : " & IIf(p_status.IO_Waiting = Omron.Omron_Remote_IO_Waiting.On, "On", "Off")).PadRight(80, " ")

            Return str_status
        End Function

        ''' <summary>
        ''' Test whether PLC is available or not
        ''' </summary>
        Public Function Test() As Boolean
            Return (Test(UnitAddress))
        End Function

        ''' <summary>
        ''' Test whether PLC is available or not
        ''' </summary>    
        Public Function Test(ByVal unit_address As Integer) As Boolean
            If Link Is Nothing Then Throw New OmronException("The object is not initilized yet!")
            If Not ISOpen Then Throw New OmronException("Port is not opened")

            Dim lng_length As Integer, tmr_indicator As Long = 0, _
            length_predicted As Integer = 11
            Dim str_in As String
            Dim str_out As String = "@" & _
                Math.Abs(unit_address Mod 32).ToString.PadLeft(2, "0"c) & "TS" & "00"
            str_out = str_out & FCSCode_Set(str_out) & "*" & vbCr
            'Fixed Length for Testing PLC
            lng_length = (length_predicted * ByteTime * 1000) + delay_ms
            lng_length *= TimeSpan.TicksPerMillisecond

            str_in = Link.ReadExisting()
            str_in = ""
            Link.Write(str_out)

            tmr_indicator = DateAndTime.Now.Ticks
            Do
                str_in = str_in & Link.ReadExisting
            Loop While ((str_in.Length < length_predicted) And ((DateAndTime.Now.Ticks - tmr_indicator) < lng_length))

            If (FCSCode_Get(str_in)) Then Me.Alive = True : Return True
            Me.Alive = False : Return False
        End Function

        ''' <summary>
        ''' Convert UShort to Hexadecimal
        ''' </summary>    
        Protected Shared Function UShort2hex(ByVal value As UShort) As String
            Return Hex(value).PadLeft(4, "0"c)
        End Function

        ''' <summary>
        ''' Convert UShortArr to Hexadecimal
        ''' </summary>    
        Protected Shared Function UShortArr2hex(ByVal arr_ushort As UShort()) As String
            Dim I As Integer = LBound(arr_ushort), u_b As Integer = UBound(arr_ushort)
            Dim str_out As String = ""
            Do
                str_out = str_out & UShort2hex(arr_ushort(I))
                I += 1
            Loop While Not (I > u_b)
            Return str_out
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As Boolean(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As Boolean(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(unit_add, arr_in, start_add, memory_area, Omron_Data_Type.Boolean))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As Byte(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As Byte(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(unit_add, arr_in, start_add, memory_area, Omron_Data_Type.Byte))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As Char(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As Char(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(unit_add, arr_in, start_add, memory_area, Omron_Data_Type.Char))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As Double(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As Double(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(unit_add, arr_in, start_add, memory_area, Omron_Data_Type.Double))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As Integer(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As Integer(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(unit_add, arr_in, start_add, memory_area, Omron_Data_Type.Integer))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As Short(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As Short(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(unit_add, arr_in, start_add, memory_area, Omron_Data_Type.Short))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As Single(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As Single(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(unit_add, arr_in, start_add, memory_area, Omron_Data_Type.Single))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As UShort(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As UShort(), ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write) As Omron_Response_Code
            Return (Write(unit_add, arr_in, start_add, memory_area, Omron_Data_Type.UShort))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal arr_in As System.Array, ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write, ByVal in_type As Omron_Data_Type) As Omron_Response_Code
            Return (Write(UnitAddress, arr_in, start_add, memory_area, in_type))
        End Function

        ''' <summary>
        ''' Write data to specific register
        ''' </summary>    
        Public Function Write(ByVal unit_add As Integer, ByVal arr_in As System.Array, ByVal start_add As Integer, ByVal memory_area As Omron_Command_Header_Write, ByVal in_type As Omron_Data_Type) As Omron_Response_Code
            If Link Is Nothing Then Throw New OmronException("The object is not initilized yet!")
            If Not ISOpen Then Throw New OmronException("Port is not opened")
            If start_add > 9999 Or start_add < 0 Then Throw New OmronException("Start Adrress is not valid!")

            Dim lng_length As Integer = 0, str_indicator As String = 1
            Dim p_response As Omron_Response_Code = 0, _
            tmr_indicator As Long = 0, length_predicted As Integer = 12
            Dim str_out As String = "", str_arr As String = ""
            Dim str_in As String = ""

            If in_type = Omron_Data_Type.Boolean Then
                str_arr = BoolArr2Hex(arr_in)
            ElseIf in_type = Omron_Data_Type.Byte Then
                str_arr = ByteArr2Hex(arr_in)
            ElseIf in_type = Omron_Data_Type.Char Then
                Dim char_arr As Char() = arr_in
                str_arr = char_arr
            ElseIf in_type = Omron_Data_Type.Double Then
                str_arr = DoubleArr2Hex(arr_in)
            ElseIf in_type = Omron_Data_Type.Integer Then
                str_arr = IntegerArr2Hex(arr_in)
            ElseIf in_type = Omron_Data_Type.Short Then
                str_arr = ShortArr2Hex(arr_in)
            ElseIf in_type = Omron_Data_Type.Single Then
                str_arr = SingleArr2Hex(arr_in)
            Else
                str_arr = UShortArr2hex(arr_in)
            End If
            If (str_arr.Length Mod 4) <> 0 Then str_arr.PadLeft(((str_arr.Length \ 4) + 1) * 4, "0"c)
            Do
                lng_length = (length_predicted * ByteTime * 1000) + (2 * delay_ms)
                lng_length *= TimeSpan.TicksPerMillisecond

                str_out = "@" & _
                    Math.Abs(unit_add Mod 32).ToString.PadLeft(2, "0"c) & _
                    str_Command_Header(memory_area) & _
                    (start_add + ((str_indicator - 1) * 0.25)).ToString.PadLeft(4, "0"c) & _
                    Mid$(str_arr, str_indicator, 116)
                str_out = str_out & FCSCode_Set(str_out) & "*" & vbCr

                str_in = Link.ReadExisting()
                str_in = ""

                Link.Write(str_out)
                'It first waits on, then reads and after all calculates the remained time for repeating this period
                tmr_indicator = DateAndTime.Now.Ticks

                Do
                    str_in = str_in & Link.ReadExisting
                Loop While ((str_in.Length < length_predicted) And ((DateAndTime.Now.Ticks - tmr_indicator) < lng_length))

                If Not (FCSCode_Get(str_in)) Then Throw New OmronException("Comminucation Error, Please check the line!")
                'First, Extracting Responce Code
                p_response = Hex2Int(Mid$(str_in, 6, 2))
                str_indicator += 116
            Loop While (Not (str_indicator > str_arr.Length))
            Return p_response
        End Function

        ''' <summary>
        ''' Set PLC Status, Ordered by User
        ''' </summary>    
        Public Function WriteStatus(ByVal p_status As Omron_Status_Mode) As Omron_Response_Code
            Return (WriteStatus(UnitAddress, p_status))
        End Function

        ''' <summary>
        ''' Set PLC Status, Ordered by User
        ''' </summary>    
        Public Function WriteStatus(ByVal unit_add As Integer, ByVal p_status As Omron_Status_Mode) As Omron_Response_Code
            If Link Is Nothing Then Throw New OmronException("The object is not initilized yet!")
            If Not ISOpen Then Throw New OmronException("Port is not opened")

            Dim lng_length As Integer, int_temp As Integer = p_status, _
            tmr_indicator As Long = 0, length_predicted As Integer = 11

            Dim str_in As String, str_temp = int_temp.ToString.PadLeft(4, "0"c)
            Dim str_out As String = "@" & _
                Math.Abs(unit_add Mod 32).ToString.PadLeft(2, "0"c) & "SC" & str_temp
            str_out = str_out & FCSCode_Set(str_out) & "*" & vbCr

            'Fixed Length for response of ReadStatus
            lng_length = (length_predicted * ByteTime * 1000) + delay_ms
            lng_length *= TimeSpan.TicksPerMillisecond

            str_in = Link.ReadExisting()
            str_in = ""
            Link.Write(str_out)

            tmr_indicator = DateAndTime.Now.Ticks
            Do
                str_in = str_in & Link.ReadExisting
            Loop While ((str_in.Length < length_predicted) And ((DateAndTime.Now.Ticks - tmr_indicator) < lng_length))

            If Not (FCSCode_Get(str_in)) Then Throw New OmronException("Comminucation Error, Please check the line!")
            'First, Extracting Responce Code
            Return (Hex2Int(Mid$(str_in, 6, 2)))

        End Function

#End Region

#Region "Properties"

        ''' <summary>
        ''' Get/Set IsAlive Status
        ''' </summary>
        Private Property Alive() As Boolean
            Get
                Return is_alive
            End Get
            Set(ByVal value As Boolean)
                is_alive = value
            End Set
        End Property

        ''' <summary>
        ''' Get Buffer Array Contents
        ''' </summary>
        Public ReadOnly Property BArr() As ArrayList
            Get
                Return inuse_String
            End Get
        End Property

        ''' <summary>
        ''' Get index'th element of Buffer Array Contents
        ''' </summary>
        Public ReadOnly Property BArr(ByVal index As Integer) As String
            Get
                If index >= inuse_String.Count Then Throw New OmronException("Index is not in range!")
                Return inuse_String(index)
            End Get
        End Property

        ''' <summary>
        ''' Get Number of Strings Added into Buffer Array
        ''' </summary>
        Public ReadOnly Property BArrCount() As Integer
            Get
                Return inuse_String.Count
            End Get
        End Property

        ''' <summary>
        ''' Get/Set Each Byte Responce Time
        ''' </summary>
        Private Property ByteTime() As Single
            Get
                Return byte_time
            End Get
            Set(ByVal value As Single)
                byte_time = value
            End Set
        End Property

        ''' <summary>
        ''' Get/Set CommPort Settings
        ''' </summary>
        Public Property CommSetting() As SerialPort
            Get
                Return Link
            End Get
            Set(ByVal value As SerialPort)
                If ISOpen Then Throw New OmronException("Assigning is prohibited while port is open!")
                Link = value
                With Link
                    Dim tmp_len As Single = 0
                    Select Case .StopBits
                        Case StopBits.None
                            tmp_len += 0
                        Case StopBits.One
                            tmp_len += 1
                        Case StopBits.OnePointFive
                            tmp_len += 1.5
                        Case StopBits.Two
                            tmp_len += 2
                    End Select
                    Select Case .Parity
                        Case Parity.Even
                            tmp_len += 1
                        Case Parity.Mark
                            tmp_len += 1
                        Case Parity.Odd
                            tmp_len += 1
                        Case Parity.Space
                            tmp_len += 1
                        Case Parity.None
                            tmp_len += 0
                    End Select
                    ByteTime = CSng(.DataBits + 1 + tmp_len) / CSng(.BaudRate)
                End With
            End Set
        End Property

        ''' <summary>
        ''' Verify whether PLC Is Alive or Not
        ''' </summary>
        Public ReadOnly Property ISAlive() As Boolean
            Get
                Return is_alive
            End Get
        End Property

        ''' <summary>
        ''' Verify whether CommPort Opened or Not
        ''' </summary>
        Public ReadOnly Property ISOpen() As Boolean
            Get
                Return is_open
            End Get
        End Property

        ''' <summary>
        ''' Get/Set IsOpen Status
        ''' </summary>
        Private Property Opened() As Boolean
            Get
                Return is_open
            End Get
            Set(ByVal value As Boolean)
                is_open = value
            End Set
        End Property

        Friend Property Path() As String
            Get
                Return p_filepath
            End Get
            Set(ByVal value As String)
                If Trim(value) <> "" Then p_filepath = value
            End Set
        End Property

        ''' <summary>
        ''' Get/Set Active PLCUnit Address
        ''' </summary>
        Public Property UnitAddress() As Integer
            Get
                Return unit_Address
            End Get
            Set(ByVal value As Integer)
                If value < 0 Or value > 31 Then Throw New OmronException("Invalid Unit Address!")
                unit_Address = value
            End Set
        End Property

#End Region

    End Class

#End Region

#End Region

#Region "Omron Enums"

    Friend Enum Omron_Command_Header
        Test = &H0
        Status_Read = &H1
        Error_Read = &H2
        IR_Area_Read = &H3
        HR_Area_Read = &H4
        AR_Area_Read = &H5
        LR_Area_Read = &H6
        TC_Status_Read = &H7
        DM_Area_Read = &H8
        FM_Index_Read = &H9
        FM_Data_Read = &HA
        PV_Read = &HB
        SV_Read_1 = &HC
        SV_Read_2 = &HD
        SV_Read_3 = &HE
        Status_Write = &HF
        IR_Area_Write = &H10
        HR_Area_Write = &H11
        AR_Area_Write = &H12
        LR_Area_Write = &H13
        TC_Status_Write = &H14
        DM_Area_Write = &H15
        FM_Area_Write = &H16
        PV_Write = &H17
        SV_Change_1 = &H18
        SV_Change_2 = &H19
        SV_Change_3 = &H1A
        Forced_Set = &H1B
        Forced_Reset = &H1C
        Multiple_Forced_Set_Reset = &H1D
        MULTIPLE_FORCED_SET_RESET_STATUS_READ = &H1E
        Forced_Set_Reset_Cancel = &H1F
        PC_Model_Read = &H20
        DM_HighSpeed_Read = &H21
        Abort_And_Initialize = &H22
        Transmit_C200Only = &H23
        Response_to_an_Undefined_Command = &H24
        Response_Indicating_an_Unprocessed_Command = &H25
        PROGRAM_READ = &H26
        IO_TABLE_READ = &H27
        PROGRAM_WRITE = &H28
        IO_TABLE_GENERATE = &H29
        IO_REGISTER = &H2A
        IO_READ = &H2B
    End Enum

    Friend Enum Omron_Command_Header_Read
        IR_Area_Read = &H3
        HR_Area_Read = &H4
        AR_Area_Read = &H5
        LR_Area_Read = &H6
        DM_Area_Read = &H8
        FM_Index_Read = &H9
        FM_Data_Read = &HA
        PV_Read = &HB
    End Enum

    Friend Enum Omron_Command_Header_Write
        IR_Area_Write = &H10
        HR_Area_Write = &H11
        AR_Area_Write = &H12
        LR_Area_Write = &H13
        DM_Area_Write = &H15
        PV_Write = &H17
    End Enum

    Friend Enum Omron_Command_SubHeader As Byte
        IO_Registers = 0
        IO_read = 1
    End Enum

    Friend Enum Omron_Duplex_System
        StandBy = 0
        Active = 1
    End Enum

    Friend Enum Omron_Error_B0_To_B3_Remote_IO_Error
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_C04_to_04_IO_Bus_Error As Byte
        [off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_Data_From_IO_Bus As Byte
        Group_1_Control_Signal_Error = 0
        Group_2_Data_Bus_Failure = 1
        Group_3_Address_Bus_Failure = 2
    End Enum

    Friend Enum Omron_Error_Duplex_Bus_Error_C2000H_Only_Intelligent_IO_Error_C200H_HS_only As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_E0_IO_Setting_Error As Byte
        [off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_E1_IO_Unit_Over As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_F0_End_Instruction_Missing
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_F1_Memory_Error As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_F2_Jump_Instruction_Error As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_F3_Program_Error
        [off] = 0
        [on] = 1
    End Enum

    Friend Enum Omron_Error_F4_RTI_Instruction_Error
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_F7_Battery_Failure As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_F7_IO_Verify_Error As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_F8_Cycle_Timer_Over
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_FAL As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_FALS As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_Diagnosis
        [Off] = 0
        InProgress = 1
    End Enum

    Friend Enum Omron_Error_Host_Link_Unit_Transmission_Error As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_PC_Link_Error As Byte
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Error_Unit_Number As Byte
        Rack_CPU = 0
        Rack_01 = 1
        Rack_02 = 2
        Rack_03 = 3
        Rack_04 = 4
        Rack_05 = 5
        Rack_06 = 6
        Rack_07 = 7
    End Enum

    Friend Enum Omron_Forced_Set_Reset
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_PLC_Mode
        Program = 0
        None = 1
        Run = 2
        Monitor = 3
    End Enum

    Friend Enum Omron_Program_Area
        ROM = 0
        RAM = 1
    End Enum

    Friend Enum Omron_Program_Area_Size
        None = 0
        _8Kbyte = 1
        _16Kbyte = 2
        _24Kbyte = 3
        _32Kbyte = 4
        _48Kbyte = 5
        _64Kbyte = 6
    End Enum

    Friend Enum Omron_Remote_IO_Waiting
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Response_Code
        Unknown = -1
        Command_Completed_Normally = 0
        Exe_Blocked_Run_Mode = 1
        Exe_Blocked_Monitor_Mode = 2
        Exe_Blocked_PROM_mounte = 3
        Address_Overflow = 4
        Exe_Blocked_Program_Mode = &HB
        Exe_Blocked_Debug_Mode = &HC
        Exe_Blocked_HotLink_is_Local = &HD
        Parity_Err = &H10
        Frame_Err = &H11
        Overrun = &H12
        FCS_Err = &H13
        Command_Format_Err = &H14
        Incorrect_Data_Area = &H15
        Invalid_Instruction = &H16
        Frame_Length_Err = &H18
        Exe_Blocked_UnExe_Err = &H19
        IO_Table_Gen_Impossible = &H20
        PC_s_CPU = &H21
        Memory_not_Exist = &H22
        Memory_Write_Protected = &H23
        Abort_Parity_Err = &HA0
        Abort_Frame_Err = &HA1
        Abort_Overrun_Err = &HA2
        Abort_FCS_Err = &HA3
        Abort_Format_Err = &HA4
        Abort_Entry_Number_Err = &HA5
        Abort_Frame_Length_Err = &HA6
        Exe_Blocked_Less_16KByte = &HB0
    End Enum

    Friend Enum Omron_Data_Type As Byte
        [Boolean] = 0
        [Byte] = 1
        [Char] = 2
        [Double] = 3
        [Integer] = 4
        [Short] = 5
        [Single] = 6
        [UShort] = 7
    End Enum

    Friend Enum Omron_StartSwitch
        [Off] = 0
        [On] = 1
    End Enum

    Friend Enum Omron_Status_Mode
        Program = 0
        Monitor = 2
        Run = 3
    End Enum

    Friend Enum Omron_TransitionType
        Simples = 0
        Duplex = 1
    End Enum

#End Region

#Region "Omron Structures"

    Friend Structure Omron_Error_LowerByte
        Dim Err_B0 As Omron_Error_B0_To_B3_Remote_IO_Error
        Dim Err_E0 As Omron_Error_E0_IO_Setting_Error
        Dim Err_E1 As Omron_Error_E1_IO_Unit_Over
        Dim Err_F7 As Omron_Error_F7_IO_Verify_Error
        Dim Err_F8 As Omron_Error_F8_Cycle_Timer_Over
    End Structure

    Friend Structure Omron_Error_MiddleByte
        Dim Err_F7_Battery As Omron_Error_F7_Battery_Failure
        Dim Err_Data_IOBus As Omron_Error_Data_From_IO_Bus
        Dim Err_Dup_IO As Omron_Error_Duplex_Bus_Error_C2000H_Only_Intelligent_IO_Error_C200H_HS_only
        Dim Err_FAL As Omron_Error_FAL
        Dim Err_UnitNo As Omron_Error_Unit_Number
    End Structure

    Friend Structure Omron_Error_UpperByte
        Dim Err_C0 As Omron_Error_C04_to_04_IO_Bus_Error
        Dim Err_F0 As Omron_Error_F0_End_Instruction_Missing
        Dim Err_F1 As Omron_Error_F1_Memory_Error
        Dim Err_F2 As Omron_Error_F2_Jump_Instruction_Error
        Dim Err_F3 As Omron_Error_F3_Program_Error
        Dim Err_F4 As Omron_Error_F4_RTI_Instruction_Error
        Dim Err_FALS As Omron_Error_FALS
        Dim Err_HostLink As Omron_Error_Host_Link_Unit_Transmission_Error
        Dim Err_PCLink As Omron_Error_PC_Link_Error
    End Structure

    Friend Structure Omron_Status
        Dim False_Instruction As Boolean
        Dim TransitionType As Omron_TransitionType
        Dim StartSwitch As Omron_StartSwitch
        Dim Error_Diagnosis As Omron_Error_Diagnosis
        Dim Forced_Set_Reset As Omron_Forced_Set_Reset
        Dim IO_Waiting As Omron_Remote_IO_Waiting
        Dim PLC_Mode As Omron_PLC_Mode
        Dim Duplex_System As Omron_Duplex_System
        Dim Program_Area As Omron_Program_Area
        Dim Program_Area_Size As Omron_Program_Area_Size
    End Structure

#End Region

End Namespace
