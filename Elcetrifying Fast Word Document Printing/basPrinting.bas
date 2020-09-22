Attribute VB_Name = "basPrinting"
   Option Explicit

   
   'Declaration for Default Printer
   Public Type PRINTER_DEFAULTS

       pDatatype As Long
       pDevmode As Long
       DesiredAccess As Long
   End Type

   'Printer Info Settings
   Public Type PRINTER_INFO_2
       pServerName As Long
       pPrinterName As Long
       pShareName As Long
       pPortName As Long
       pDriverName As Long
       pComment As Long
       pLocation As Long
       pDevmode As Long       ' Pointer to DEVMODE
       pSepFile As Long
       pPrintProcessor As Long
       pDatatype As Long
       pParameters As Long
       pSecurityDescriptor As Long  ' Pointer to SECURITY_DESCRIPTOR
       Attributes As Long


       Priority As Long
       DefaultPriority As Long
       StartTime As Long
       UntilTime As Long
       Status As Long
       cJobs As Long
       AveragePPM As Long
   End Type

   'Dev Mode Settings
   Public Type DEVMODE
       dmDeviceName As String * 32

       dmSpecVersion As Integer
       dmDriverVersion As Integer
       dmSize As Integer
       dmDriverExtra As Integer
       dmFields As Long
       dmOrientation As Integer
       dmPaperSize As Integer
       dmPaperLength As Integer
       dmPaperWidth As Integer
       dmScale As Integer
       dmCopies As Integer
       dmDefaultSource As Integer
       dmPrintQuality As Integer
       dmColor As Integer
       dmDuplex As Integer
       dmYResolution As Integer
       dmTTOption As Integer
       dmCollate As Integer
       dmFormName As String * 32
       dmUnusedPadding As Integer
       dmBitsPerPel As Integer
       dmPelsWidth As Long
       dmPelsHeight As Long
       dmDisplayFlags As Long
       dmDisplayFrequency As Long
       dmICMMethod As Long
       dmICMIntent As Long
       dmMediaType As Long
       dmDitherType As Long
       dmReserved1 As Long
       dmReserved2 As Long
   End Type

   'Constant Declaration for Printing Option
   Public Const DM_DUPLEX = &H1000&
   Public Const DM_IN_BUFFER = 8

   Public Const DM_OUT_BUFFER = 2
   Public Const PRINTER_ACCESS_ADMINISTER = &H4
   Public Const PRINTER_ACCESS_USE = &H8
   Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
   Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
             PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

   'API Declaration for Printing
   Public Declare Function ClosePrinter Lib "winspool.drv" _
    (ByVal hPrinter As Long) As Long
   Public Declare Function DocumentProperties Lib "winspool.drv" _
     Alias "DocumentPropertiesA" (ByVal hwnd As Long, _
     ByVal hPrinter As Long, ByVal pDeviceName As String, _
     ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, _
     ByVal fMode As Long) As Long
   Public Declare Function GetPrinter Lib "winspool.drv" Alias _
     "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
     pPrinter As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
   Public Declare Function OpenPrinter Lib "winspool.drv" Alias _
     "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
     pDefault As PRINTER_DEFAULTS) As Long
   Public Declare Function SetPrinter Lib "winspool.drv" Alias _
     "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
     pPrinter As Byte, ByVal Command As Long) As Long

   Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (pDest As Any, pSource As Any, ByVal cbLength As Long)
 
   ' ==================================================================
   ' SetPrinterDuplex
   '
   '  Programmatically set the Duplex flag for the specified printer
   '  driver's default properties.
   '
   '  Returns: True on success, False on error. (An error will also

   '  display a message box. This is done for informational value
   '  only. You should modify the code to support better error
   '  handling in your production application.)
   '
   '  Parameters:
   '    sPrinterName - The name of the printer to be used.
   '
   '    nDuplexSetting - One of the following standard settings:
   '       1 = None
   '       2 = Duplex on long edge (book)
   '       3 = Duplex on short edge (legal)
   '
   ' ==================================================================
   Public Function SetPrinterDuplex(ByVal sPrinterName As String, _
       ByVal nDuplexSetting As Long) As Boolean

      Dim hPrinter As Long
      Dim pd As PRINTER_DEFAULTS
      Dim pinfo As PRINTER_INFO_2
      Dim dm As DEVMODE
   
      Dim yDevModeData() As Byte
      Dim yPInfoMemory() As Byte
      Dim nBytesNeeded As Long
      Dim nRet As Long, nJunk As Long
   
      On Error GoTo cleanup
   
      'Checking Duplexing Mode
      If (nDuplexSetting < 1) Or (nDuplexSetting > 3) Then
         MsgBox "Error: dwDuplexSetting is incorrect."
         Exit Function
      End If
      
      'Configure and Open the Printer
      pd.DesiredAccess = PRINTER_ALL_ACCESS
      nRet = OpenPrinter(sPrinterName, hPrinter, pd)
      If (nRet = 0) Or (hPrinter = 0) Then
         If Err.LastDllError = 5 Then
            MsgBox "Access denied -- See the article for more info."
         Else
            MsgBox "Cannot open the printer specified " & _
              "(make sure the printer name is correct)."
         End If
         Exit Function
      End If
        
      'Assign Document Properties with Printer Handle and Name
      nRet = DocumentProperties(0, hPrinter, sPrinterName, 0, 0, 0)
      If (nRet < 0) Then
         MsgBox "Cannot get the size of the DEVMODE structure."
         GoTo cleanup
      End If
      ReDim yDevModeData(nRet + 100) As Byte
      nRet = DocumentProperties(0, hPrinter, sPrinterName, _
                  VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)
      If (nRet < 0) Then
         MsgBox "Cannot get the DEVMODE structure."
         GoTo cleanup
      End If
        
      'Copy Memory
      Call CopyMemory(dm, yDevModeData(0), Len(dm))
      'Check Modification of Duplex option for Printer
      If Not CBool(dm.dmFields And DM_DUPLEX) Then
        MsgBox "You cannot modify the duplex flag for this printer " & _
               "because it does not support duplex or the driver " & _
               "does not support setting it from the Windows API."
         GoTo cleanup
      End If
        
      dm.dmDuplex = nDuplexSetting
      Call CopyMemory(yDevModeData(0), dm, Len(dm))
        
      'Check Availabilty of Duplex Printing
      nRet = DocumentProperties(0, hPrinter, sPrinterName, _
        VarPtr(yDevModeData(0)), VarPtr(yDevModeData(0)), _
        DM_IN_BUFFER Or DM_OUT_BUFFER)
        
      If (nRet < 0) Then
        MsgBox "Unable to set duplex setting to this printer."
        GoTo cleanup
      End If
        
      Call GetPrinter(hPrinter, 2, 0, 0, nBytesNeeded)
      If (nBytesNeeded = 0) Then GoTo cleanup
        
      ReDim yPInfoMemory(nBytesNeeded + 100) As Byte
        
      'Get printer and load into memory
      nRet = GetPrinter(hPrinter, 2, yPInfoMemory(0), nBytesNeeded, nJunk)
      If (nRet = 0) Then
         MsgBox "Unable to get shared printer settings."
         GoTo cleanup
      End If
        
      'Copy document into memory
      Call CopyMemory(pinfo, yPInfoMemory(0), Len(pinfo))
      pinfo.pDevmode = VarPtr(yDevModeData(0))
      pinfo.pSecurityDescriptor = 0
      Call CopyMemory(yPInfoMemory(0), pinfo, Len(pinfo))
        
      'Set printer for Printing
      nRet = SetPrinter(hPrinter, 2, yPInfoMemory(0), 0)
      If (nRet = 0) Then
         MsgBox "Unable to set shared printer settings."
      End If
        
      'Result of printing success
      SetPrinterDuplex = CBool(nRet)

cleanup:
      If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)

   End Function
