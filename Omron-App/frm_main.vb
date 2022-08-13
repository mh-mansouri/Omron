
#Region "Refrerences"

#End Region

Friend Class frm_main

#Region "Fields"

    Private PLC As Omron.OmronPLC

#End Region

#Region "Events"

#Region "Form Events"

    Private Sub frm_main_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If Not PLC Is Nothing Then
            If PLC.ISOpen Then PLC.Close()
            PLC = Nothing
        End If
    End Sub

    Private Sub frm_main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        PLC = New Omron.OmronPLC(New System.IO.Ports.SerialPort("COM1", 115200, IO.Ports.Parity.Even, 8, IO.Ports.StopBits.One))
        PLC.Open()
    End Sub

#End Region

#Region "Command Button"

    Private Sub cmd_clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_clear.Click
        txt_read.Text = ""
    End Sub

    Private Sub cmd_PLCread_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_PLCread.Click
        Dim p_response As Omron.Omron_Response_Code
        Dim reg_start As Integer = 90, reg_count As Integer = 30
        PLC.BArr_ReadAdd(Omron.Omron_Command_Header_Read.DM_Area_Read, reg_start, reg_count)
        Dim arr_read As Char() = New Char() {}
        Dim lng_length As Long = 0
        lng_length = DateAndTime.Now.Ticks
        arr_read = PLC.Read(PLC.BArrCount - 1, p_response)
        lng_length = DateAndTime.Now.Ticks - lng_length
        Dim str_read As String = arr_read
        Dim str_out As String = ""
        txt_read.Text = ""
        Try
            For I As Integer = 0 To reg_count - 1
                str_out = str_out & "DM ( " & (reg_start + I).ToString().PadLeft(4, "0"c) & " ) : " & str_read.Substring(I * 4, 4) & vbNewLine
            Next
            str_out = str_out & lng_length / TimeSpan.TicksPerMillisecond & " milliSec"
        Catch ex As Exception
            str_out = "No datat is available"
        End Try
        txt_read.Text = str_out
    End Sub

    Private Sub cmd_PLCStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_PLCStatus.Click
        Dim p_status As Omron.Omron_Status
        Dim p_response As Omron.Omron_Response_Code
        Dim p_Ubyte As Omron.Omron_Error_UpperByte
        Dim p_Mbyte As Omron.Omron_Error_MiddleByte
        Dim p_Lbyte As Omron.Omron_Error_LowerByte

        lst_read.Items.Clear()

        'First Writing PLC is Available
        lst_read.Items.Add("PLC is Avalilable".PadRight(80, "-"))
        'Writing PLC Responce Message
        lst_read.Items.Add(("PLC Status Respose: " & PLC.ReadStatus(p_response, p_status)).PadRight(80, " "))

        'Writing PLC Status
        lst_read.Items.Add("PLC Status".PadRight(80, "-"))
        Dim str_status As String() = PLC_Status(p_status)
        For Each str_resp As String In str_status
            lst_read.Items.Add(str_resp)
        Next

        'Writing PLC responce Code
        lst_read.Items.Add("PLC Response Code".PadRight(80, "-"))
        Dim str_response As String() = PLC_Response(p_response)
        For Each str_resp As String In str_response
            lst_read.Items.Add(str_resp)
        Next

        'Writing PLC Error Code
        lst_read.Items.Add("PLC Error Code : ".PadRight(80, "-"))
        Dim str_err_temp As String = PLC.ReadError(p_response, p_Ubyte, p_Mbyte, p_Lbyte, False)
        Dim str_err As String() = PLC_Error(p_response, p_Ubyte, p_Mbyte, p_Lbyte)
        For Each str_error As String In str_err
            lst_read.Items.Add(str_error)
        Next
        lst_read.Items.Add(str_err_temp.PadRight(80, " "))
    End Sub

    Private Sub cmd_PLCTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_PLCTest.Click
        lst_read.Items.Clear()

        If PLC.Test Then
            lst_read.Items.Add("PLC is Avalilable")
            cmd_PLCStatus.Enabled = True
            cmd_PLCread.Enabled = True
            cmd_write.Enabled = True
        Else
            lst_read.Items.Add("PLC is Not Avalilable")
            cmd_PLCStatus.Enabled = False
            cmd_PLCread.Enabled = False
            cmd_write.Enabled = False
        End If
    End Sub

    Private Sub cmd_write_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_write.Click
        Dim lng_time As Long = DateAndTime.Now.Ticks
        'Dim p_respose As Omron.Omron_Response_Code = PLC.Write(New Char() {"F", "F", "F", "F"}, 100, Omron.Omron_Command_Header_Write.DM_Area_Write, Omron.Omron_Data_Type.Char)
        Dim p_respose As Omron.Omron_Response_Code = PLC.Write(New Char() {"0", "0", "0", "0"}, 101, Omron.Omron_Command_Header_Write.DM_Area_Write)
        lng_time = DateAndTime.Now.Ticks - lng_time
        lst_read.Items.Clear()
        lst_read.Items.Add("Writing one register took : " & lng_time / TimeSpan.TicksPerMillisecond & " milliSec")
    End Sub

#End Region

#End Region

#Region "Funs and Subs"

    Private Function PLC_Error(ByVal p_response As Omron.Omron_Response_Code, ByVal uByte As Omron.Omron_Error_UpperByte, ByVal mByte As Omron.Omron_Error_MiddleByte, ByVal lByte As Omron.Omron_Error_LowerByte) As String()
        Dim str_err As String() = New String(19) {}
        str_err(0) = PLC_Response(p_response)(0)

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

    Private Function PLC_Response(ByVal p_responce As Omron.Omron_Response_Code) As String()
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

    Private Function PLC_Status(ByVal p_status As Omron.Omron_Status) As String()
        Dim str_status As String() = Nothing
        str_status = New String(4) {}


        str_status(0) = ("System Duplex System Status : " & IIf(p_status.Duplex_System = Omron.Omron_Duplex_System.Active, " Actived", "Standby")).PadRight(80, " ")
        str_status(1) = ("System Error Diagnosis Status : " & IIf(p_status.Error_Diagnosis = Omron.Omron_Error_Diagnosis.InProgress, "InProgress", "Off")).PadRight(80, " ")
        str_status(2) = ("System Instruction Status : " & IIf(p_status.False_Instruction, "False", "Correct")).PadRight(80, " ")
        str_status(3) = ("System Force Set/Reset Status : " & IIf(p_status.Forced_Set_Reset = Omron.Omron_Forced_Set_Reset.On, "On", "Off")).PadRight(80, " ")
        str_status(4) = ("System Remote IO Waiting Status : " & IIf(p_status.IO_Waiting = Omron.Omron_Remote_IO_Waiting.On, "On", "Off")).PadRight(80, " ")

        Return str_status
    End Function

#End Region

End Class
