Option Explicit On

Module VBAutomation

    Dim AspenPlus As HappLS

    Sub Main()
        Dim ihAPsim As IHapp
        ihAPsim = OpenSimulation()

        'Call GetCollectionExample(ihAPsim)
        'Call GetScalarValuesExample(ihAPsim)
        'Call ListBlocksExample(ihAPsim)
        'Call UnitStringExample(ihAPsim)
        'Call UnitsConversionExample(ihAPsim)
        'Call UnitsChangeExample(ihAPsim)
        'Call TempProfExample(ihAPsim)
        'Call CompProfExample(ihAPsim)
        'Call ReacCoeffExample(ihAPsim)
        Call ConnectivityExample(ihAPsim)
        Call RunExample(ihAPsim)
        'Call CloseSimulation(ihAPsim)
    End Sub

    Function OpenSimulation() As IHapp

        Dim ihAPsim As IHapp
        On Error GoTo ErrorHandler
        Dim VERSION As String = "V8.4"
        Dim VERSIONNUMBER As String = "30.0"

        ' define the path to the AspenPlus examples folder
        Dim defaultpath As String
        If (8 = IntPtr.Size Or Not String.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"))) Then
            defaultpath = Environment.GetEnvironmentVariable("ProgramFiles(x86)") + "\AspenTech\Aspen Plus " + VERSION + "\GUI"
        Else
            defaultpath = Environment.GetEnvironmentVariable("ProgramFiles") + "\AspenTech\Aspen Plus " + VERSION + "\GUI"
        End If

        Dim path As String
        Dim regKey As Microsoft.Win32.RegistryKey
        regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\AspenTech\Aspen Plus\" + VERSIONNUMBER + "\mm", False)
        If (regKey Is Nothing) Then
            regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Wow6432Node\AspenTech\Aspen Plus\" + VERSIONNUMBER + "\mm", False)
        End If
        If (Not regKey Is Nothing) Then
            path = regKey.GetValue("mmtop", defaultpath).ToString()
            regKey.Close()
            path = path + "\Examples\"
        Else
            path = defaultpath + "\Examples\"
        End If

        ' open existing simulation
        AspenPlus = GetObject(path & "pfdtut.bkp")
        ihAPsim = AspenPlus.Application

        ' display the GUI
        ihAPsim.Visible = True
        ' run the simulation
        ihAPsim.Run()

        ' return the Happ object for the AspenPlus simulator
        OpenSimulation = ihAPsim
        Exit Function
ErrorHandler:
        MsgBox("OpenSimulation raised error " & Err.Description)
        End
    End Function

    Sub GetCollectionExample(ByVal ihAPsim As IHapp)
        ' This example illustrates use of a collection object
        Dim ihRoot As IHNode
        Dim ihcolOffspring As IHNodeCol
        Dim ihOffspring As IHNode
        Dim strOut As String
        On Error GoTo ErrorHandler
        'get the root of the tree
        ihRoot = ihAPsim.Tree
        'now get the collection of nodes immediately below the Root
        ihcolOffspring = ihRoot.Elements
        strOut = "
"
        For Each ihOffspring In ihcolOffspring
            strOut = strOut & Chr(13) & ihOffspring.Name
        Next
        MsgBox("Offspring nodes are: " & strOut, , "GetCollectionExample")
        Exit Sub
ErrorHandler:
        MsgBox("GetCollectionExample raised error" & Err.Description)
    End Sub

    Sub GetScalarValuesExample(ByVal ihAPsim As IHapp)
        ' This example retrieves scalar variables from a block
        Dim ihColumn As IHNode
        Dim nStages As Long
        Dim buratio As Double
        On Error GoTo ErrorHandler
        ' navigate the tree to the RADFRAC block
        ihColumn = ihAPsim.Tree.Data.Blocks.B6
        ' get the number of stages
        nStages = ihColumn.Input.Elements("NSTAGE").Value
        ' get the boilup ratio
        buratio = ihColumn.Output.Elements("BU_RATIO").Value
        MsgBox("Number of Stages is: " & nStages & Chr(13) _
              & "Boilup Ratio is: " & buratio, , "GetScalarValuesExample")
        Exit Sub
ErrorHandler:
        MsgBox("GetScalarValuesExample raised error" & Err.Description)
    End Sub


    Sub ListBlocksExample(ByVal ihAPsim As IHapp)
        ' This example retrieves a list of blocks and their attributes
        Dim ihBlockList As IHNodeCol
        Dim ihBlock As IHNode
        Dim strOut As String
        On Error GoTo ErrorHandler

        ihBlockList = ihAPsim.Tree.Data.Blocks.Elements
        strOut = "Block" & Chr(9) & "Block Type" _
                 & Chr(9) & "Section  " & Chr(9) & "Results status"
        For Each ihBlock In ihBlockList
            strOut = strOut & Chr(13) & ihBlock.Name & Chr(9) &
                ihBlock.AttributeValue(Happ.HAPAttributeNumber.HAP_RECORDTYPE) & "  " & Chr(9) &
                ihBlock.AttributeValue(Happ.HAPAttributeNumber.HAP_SECTION) & Chr(9) &
                Status(ihBlock.AttributeValue(Happ.HAPAttributeNumber.HAP_COMPSTATUS))
        Next ihBlock
        MsgBox(strOut, , "ListBlocksExample")
        Exit Sub
ErrorHandler:
        MsgBox("ListBlocksExample raised error" & Err.Description)
    End Sub

    Function Status(ByVal CompStat As Integer) As String
        ' This function interprets a status variable and returns a string

        If ((CompStat And Happ.HAPCompStatusCode.HAP_RESULTS_SUCCESS) = Happ.HAPCompStatusCode.HAP_RESULTS_SUCCESS) Then
            Status = "Success"
        ElseIf ((CompStat And Happ.HAPCompStatusCode.HAP_RESULTS_ERRORS) = Happ.HAPCompStatusCode.HAP_RESULTS_ERRORS) Then
            Status = "Errors"
        ElseIf ((CompStat And Happ.HAPCompStatusCode.HAP_RESULTS_WARNINGS) = Happ.HAPCompStatusCode.HAP_RESULTS_WARNINGS) Then
            Status = "Warnings"
        ElseIf ((CompStat And Happ.HAPCompStatusCode.HAP_NORESULTS) = Happ.HAPCompStatusCode.HAP_NORESULTS) Then
            Status = "No results"
        ElseIf ((CompStat And Happ.HAPCompStatusCode.HAP_RESULTS_INCOMPAT) = Happ.HAPCompStatusCode.HAP_RESULTS_INCOMPAT) Then
            Status = "Incompatible with input"
        ElseIf ((CompStat And Happ.HAPCompStatusCode.HAP_RESULTS_INACCESS) = Happ.HAPCompStatusCode.HAP_RESULTS_INACCESS) Then
            Status = "In access"
        End If
    End Function

    Sub UnitStringExample(ByVal ihAPsim As IHapp)
        ' This example retrieves the units of measurement symbol for a
        ' variable
        Dim ihPresNode As IHNode
        On Error GoTo ErrorHandler
        ihPresNode = ihAPsim.Tree.Data.Blocks.B3.Output.B_PRES
        MsgBox("Flash pressure is: " & ihPresNode.Value & Chr(9) & _
                 ihPresNode.UnitString, , "UnitStringExample")
        Exit Sub
ErrorHandler:
        MsgBox("UnitStringExample raised error " & Err.Description)
    End Sub
    Sub UnitsConversionExample(ByVal ihAPsim As IHapp)
        ' This example retrieves a value both in the display units (psi) and
        ' alternative units (atm)
        Dim ihPres As IHNode
        Dim nRow As Long
        Dim nCol As Long
        Dim strDisplayUnits As String
        Dim strConvertedUnits As String
        On Error GoTo ErrorHandler
        ihPres = ihAPsim.Tree.Data.Blocks.B3.Output.B_PRES
        ' retrieve the attributes for the display units (psi)
        nRow = ihPres.AttributeValue(Happ.HAPAttributeNumber.HAP_UNITROW)
        nCol = ihPres.AttributeValue(Happ.HAPAttributeNumber.HAP_UNITCOL)
        strDisplayUnits = UnitsString(ihAPsim, nRow, nCol)
        'select the alternative unit table column (atm)
        nCol = 3
        strConvertedUnits = UnitsString(ihAPsim, nRow, nCol)
        MsgBox("Pressure in Display units: " & ihPres.Value & _
               " " & strDisplayUnits & Chr(13) & _
               "Pressure in Converted units: " & _
               ihPres.ValueForUnit(nRow, nCol) & " " & strConvertedUnits, _
               , "UnitsConversionExample")
        Exit Sub
ErrorHandler:
        MsgBox("UnitsConversionExample raised error " & Err.Description)
    End Sub
    Public Function UnitsString(ByVal ihAPsim As IHapp, ByVal nRow As Long, ByVal nCol As Long)
        ' This function returns the units of measurement symbol given
        ' the unit table row and column
        On Error GoTo UnitsStringFailed
        UnitsString = ihAPsim.Tree.Elements("Unit Table"). _
                      Elements(nRow - 1).Elements.Label(0, nCol - 1)
        Exit Function
UnitsStringFailed:
        UnitsString = ""
    End Function

    Sub UnitsChangeExample(ByVal ihAPsim As IHapp)
        ' This example shows changing the units of measurement of a variable
        Dim ihPres As IHNode
        On Error GoTo ErrorHandler
        ihPres = ihAPsim.Tree.Data.Blocks.B3.Output.B_PRES
        MsgBox("Pressure in default units: " _
                      & ihPres.Value _
                      & Chr(9) & ihPres.UnitString)
        ' change units of measurement to bar
        ihPres.AttributeValue(Happ.HAPAttributeNumber.HAP_UNITCOL, True) = 5
        MsgBox("Pressure in selected units: " _
                      & ihPres.Value _
                      & Chr(9) & ihPres.UnitString)
        Exit Sub
ErrorHandler:
        MsgBox("UnitsChangeExample raised error " & Err.Description)
    End Sub


    Sub TempProfExample(ByVal ihAPsim As IHapp)
        ' This example retrieves values for a non-scalar variable with
        ' one identifier
        Dim ihTVar As IHNode
        Dim ihStage As IHNode
        Dim strOut As String
        On Error GoTo ErrorHandler
        ihTVar = ihAPsim.Tree.Data.Blocks.B6.Output.B_TEMP
        strOut = ihTVar.Elements.DimensionName(0) & Chr(9) & ihTVar.Name
        For Each ihStage In ihTVar.Elements
            strOut = strOut & Chr(13) & ihStage.Name _
            & Chr(9) & Format(ihStage.Value, "###.00") _
            & Chr(9) & ihStage.UnitString
        Next ihStage
        MsgBox(strOut, , "TempProfExample")
        Exit Sub
ErrorHandler:
        MsgBox("TempProfExample raised error " & Err.Description)
    End Sub

    Sub CompProfExample(ByVal ihAPsim As IHapp)
        ' This example retrieves values for a non-scalar variable with
        ' two identifiers
        Dim ihTrayNode As IHNode
        Dim ihXNode As IHNode
        Dim ihCompNode As IHNode
        Dim strOut As String
        Dim nLines As Integer
        On Error GoTo ErrorHandler
        ihXNode = ihAPsim.Tree.Data.Blocks.B6.Output.Elements("X")
        nLines = 0
        For Each ihTrayNode In ihXNode.Elements
            For Each ihCompNode In ihTrayNode.Elements
                strOut = strOut & Chr(13) & ihTrayNode.Name & _
                         Chr(9) & ihCompNode.Name & Chr(9) & _
                         ihCompNode.Value
                nLines = nLines + 1
                If nLines = 40 Then
                    MsgBox(strOut, , "CompProfExample")
                    strOut = ""
                    nLines = 0
                End If
            Next ihCompNode
        Next ihTrayNode
        If nLines > 0 Then
            MsgBox(strOut, , "CompProfExample")
        End If
        Exit Sub
ErrorHandler:
        MsgBox("CompProfExample raised error " & Err.Description)
    End Sub

    Sub ReacCoeffExample(ByVal ihAPsim As IHapp)
        ' This example retrieves values for a non-scalar variable with
        ' three identifiers
        Dim ihReacNode As IHNode
        Dim ihCoeffNode As IHNode
        Dim intOff As Long
        Dim strHeading As String
        Dim strTable As String
        Dim nReacCoeff As Integer
        On Error GoTo ErrorHandler
        ihCoeffNode = ihAPsim.Tree.Data.Blocks.B2.Input.COEF
        ' loop through reaction nodes
        For Each ihReacNode In ihCoeffNode.Elements
            strHeading = ihCoeffNode.Elements.DimensionName(0) _
              & Chr(9) & ihReacNode.Elements.DimensionName(0) _
              & Chr(9) & ihReacNode.Elements.DimensionName(1)
            nReacCoeff = ihReacNode.Elements.RowCount(0)
            ' loop through coefficient nodes retrieving component and substream
            ' identifiers and coefficient values
            For intOff = 0 To nReacCoeff - 1
                strTable = strTable & Chr(13) & ihReacNode.Name & Chr(9) _
                & Chr(9) & ihReacNode.Elements.Label(0, intOff) & Chr(9) _
                & Chr(9) & ihReacNode.Elements.Label(1, intOff) & Chr(9) _
                & Chr(9) & ihReacNode.Elements.Item(intOff, intOff).Value
            Next intOff
            MsgBox(strHeading & strTable, , "ReacCoeffExample")
        Next ihReacNode
        Exit Sub
ErrorHandler:
        MsgBox("ReacCoeffExample raised error " & Err.Description)
    End Sub

    Sub ConnectivityExample(ByVal ihAPsim As IHapp)
        ' This example displays a table showing flowsheet connectivity
        Dim ihStreamList As IHNode
        Dim ihBlockList As IHNode
        Dim ihDestBlock As IHNode
        Dim ihSourceBlock As IHNode
        Dim ihStream As IHNode
        Dim strHeading As String
        Dim strTable As String
        Dim strDestBlock As String
        Dim strDestPort As String
        Dim strSourceBlock As String
        Dim strSourcePort As String
        Dim strStreamName As String
        Dim strStreamType As String

        On Error GoTo ErrorHandler
        ihStreamList = ihAPsim.Tree.Data.Streams
        ihBlockList = ihAPsim.Tree.Data.Blocks

        strHeading = "Stream" & Chr(9) & "From" _
                & Chr(9) & Chr(9) & Chr(9) & "To" & Chr(13)

        For Each ihStream In ihStreamList.Elements
            strStreamName = ihStream.Name
            strStreamType = ihStream.AttributeValue(Happ.HAPAttributeNumber.HAP_RECORDTYPE)
            ' get the destination connections
            ihDestBlock = ihStream.Elements("Ports").Elements("DEST")
            If (ihDestBlock.Elements.RowCount(0) > 0) Then
                ' there is a destination port
                strDestBlock = ihDestBlock.Elements(0).Value
                strDestPort = ihBlockList.Elements(strDestBlock). _
                Connections.Elements(strStreamName).Value
            Else
                ' it's a flowsheet product
                strDestBlock = ""
                strDestPort = ""
            End If
            ' get the source connections
            ihSourceBlock = ihStream.Elements("Ports").Elements("SOURCE")
            If (ihSourceBlock.Elements.RowCount(0) > 0) Then
                ' there is a source port
                strSourceBlock = ihSourceBlock.Elements(0).Value
                strSourcePort = ihBlockList.Elements(strSourceBlock). _
                Connections.Elements(strStreamName).Value
            Else
                ' it's a flowsheet feed
                strSourceBlock = ""
                strSourcePort = ""
            End If

            strTable = strTable & Chr(13) & strStreamName _
                      & Chr(9) & strSourceBlock _
                      & Chr(9) & strSourcePort & Chr(9) _
                      & Chr(9) & strDestBlock & Chr(9) _
                      & strDestPort
        Next ihStream
        MsgBox(strHeading & strTable, , "ConnectivityExample")
        Exit Sub

ErrorHandler:
        MsgBox("ConnectivityExample raised error" & Err.Description)

    End Sub

    Sub RunExample(ByVal ihAPsim As IHapp)
        ' This example changes a simulation parameter and re-runs the simulation
        Dim ihEngine As IHAPEngine
        Dim nStages As Object
        Dim strPrompt As String
        On Error GoTo ErrorHandler
        ihEngine = ihAPsim.Engine
EditSimulation:
        nStages = ihAPsim.Tree.Data.Blocks.B6.Input.Elements("NSTAGE").Value
        strPrompt = "Existing number of stages for column B6 = " + nStages.ToString() _
         + Chr(10) + "Enter new value for number of stages." _
         + Chr(10) + "Click 'Cancel' to exit."
        nStages = InputBox(strPrompt)
        If (nStages = "") Then GoTo finish
        ' edit the simulation
        ihAPsim.Tree.Data.Blocks.B6.Input.Elements("NSTAGE").Value = nStages
        ' run the simulation
        ihAPsim.Run()
        ' look at the status and results
        Call ListBlocksExample(ihAPsim)
        Call TempProfExample(ihAPsim)
        GoTo EditSimulation
finish:
        Exit Sub
ErrorHandler:
        MsgBox("RunExample failed with error " & Err.Description)
    End Sub

    Sub CloseSimulation(ByVal ihAPsim As IHapp)
        On Error GoTo ErrorHandler
        'Quit without saving
        ihAPsim.Quit()
        Exit Sub
ErrorHandler:
        MsgBox("CloseSimulation raised error " & Err.Description)
        End
    End Sub
End Module

