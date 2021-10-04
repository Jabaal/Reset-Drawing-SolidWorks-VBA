'**********************
'Parts of this macro are from:
'Copyright(C) 2020 Xarial Pty Limited
'Reference: https://www.codestack.net/solidworks-api/document/drawing/replace-sheet-format/
'License: https://www.codestack.net/license/
'
'Deletes all layers and adds defined standard ones
'switches out the overall drafting standard
'sets all notes to use the default font from the drafting standard
'replaces sheet format
'**********************

Const REMOVE_MODIFIED_NOTES As Boolean = True
Const FILTER_ANY As String = "*"

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Dim REPLACE_MAP As Variant
Dim stdLayerMap As Variant
Dim sPath As String


Sub main()
    'enter your desired sheet format here
    'please visit https://www.codestack.net/solidworks-api/document/drawing/replace-sheet-format/ for more details
    REPLACE_MAP = Array("*|*|C:\Solidworks\your sheet format.slddrt")
    ' this layer contains the information for all the standard layers:
    ' the following parameters are set:
    ' [Name],[DescIn],[ColorIn],[StyleIn],[WidthIn]
    ' Layer name, Description (can be empty), Color, Style as in swLineStyles_e, Width as in swLineWeights_e
    stdLayerMap = Array( _
                        Array("ANNOTATIONS", "", 0, 0, 0), _
                        Array("BEND NOTES", "", 0, 0, 0), _
                        Array("BORDERS", "", 0, 0, 0), _
                        Array("CENTERLINES", "", 0, 4, 0), _
                        Array("DIMENSIONS", "", 0, 0, 0), _
                        Array("TABLES", "", 0, 0, 0), _
                        Array("VIEWS", "", 0, 0, 0), _
                        Array("SKETCHES", "", 0, 0, 0), _
                        Array("CENTER MARKS", "", 0, 0, 0))
' set the path to drafting standards here
    sPath = "C:\Solidworks\your-drafting-standard.sldstd"

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    'Start performance enhancers
    StartNoScreenUpdate swModel
    StartLockFMT swModel
    
    
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swDraw = swApp.ActiveDoc
    
    Dim vSheetNames As Variant
    vSheetNames = swDraw.GetSheetNames
    
    Dim i As Integer
    
    Dim activeSheet As String
    activeSheet = swDraw.GetCurrentSheet().GetName
' this deletes all layers and adds the standard ones
    DeleteAndAddLayers stdLayerMap, swDraw
' this replaces the overall drawing drafting standard
    ReplaceDraftingStandard swDraw, sPath
' traverse through all sheets
    For i = 0 To UBound(vSheetNames)
        
        Dim sheetName As String
        sheetName = CStr(vSheetNames(i))
        
        Dim swSheet As SldWorks.sheet
        Set swSheet = swDraw.sheet(sheetName)
' set all notes on this drawing sheet to the default font
        SetAnnotationsToDefaultFont swDraw
        
        Dim targetSheetFormatFileName As String
        targetSheetFormatFileName = GetReplaceSheetFormat(swSheet)
'        Debug.Print targetSheetFormatFileName
        
        swDraw.ActivateSheet sheetName
' this replaces the sheet format
        ReplaceSheetFormat swDraw, swSheet, targetSheetFormatFileName
' this reloads the sheet format in case it did not work our properly and gives a message
        ReloadSheetFormat swSheet
    Next
    
    swDraw.ActivateSheet activeSheet
    ' End performance enhancers
    EndNoScreenUpdate swModel
    EndLockFMT swModel
    ' force rebuild everything
' tb implemented
    
End Sub

Function GetReplaceSheetFormat(sheet As SldWorks.sheet) As String
    
    Dim curTemplateName As String
    curTemplateName = sheet.GetTemplateName()
    
    Dim curSize As Integer
    curSize = sheet.GetSize(-1, -1)
    
    Dim i As Integer
    
    For i = 0 To UBound(REPLACE_MAP)
        
        Dim map As String
        map = REPLACE_MAP(i)
        
        Dim mapParams As Variant
        mapParams = Split(map, "|")
        
        Dim mapPaperSize As Integer
        Dim srcTemplateName As String
        
        If Trim(mapParams(0)) <> FILTER_ANY Then
            mapPaperSize = CInt(Trim(mapParams(0)))
        Else
            mapPaperSize = -1
        End If
        
        If Trim(mapParams(1)) <> FILTER_ANY Then
            srcTemplateName = CStr(Trim(mapParams(1)))
        Else
            srcTemplateName = ""
        End If
        
        If (mapPaperSize = -1 Or mapPaperSize = curSize) And (srcTemplateName = "" Or LCase(srcTemplateName) = LCase(curTemplateName)) Then
            
            Dim targetTemplateName As String

            targetTemplateName = CStr(Trim(mapParams(2)))
        
            If targetTemplateName = "" Then
                Err.Raise vbError, "", "Target template is not specified"
            End If
        
            GetReplaceSheetFormat = targetTemplateName
            Exit Function
            
        End If
        
    Next
    
    Err.Raise vbError, "", "Failed find the sheet format mathing current sheet"
    
End Function

Sub ReplaceSheetFormat(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet, targetSheetFormatFile As String)
    
    Debug.Print "Replacing '" & sheet.GetName() & "' with '" & targetSheetFormatFile & "'"
    
    Dim vProps As Variant
    vProps = sheet.GetProperties()
    
    Dim paperSize As Integer
    Dim templateType As Integer
    Dim scale1 As Double
    Dim scale2 As Double
    Dim firstAngle As Boolean
    Dim width As Double
    Dim height As Double
    Dim custPrpView As String
    
    paperSize = CInt(vProps(0))
    templateType = CInt(vProps(1))
    scale1 = CDbl(vProps(2))
    scale2 = CDbl(vProps(3))
    firstAngle = CBool(vProps(4))
    width = CDbl(vProps(5))
    height = CDbl(vProps(6))
    custPrpView = sheet.CustomPropertyView
    
    If False = draw.SetupSheet5(sheet.GetName(), paperSize, templateType, scale1, scale2, firstAngle, targetSheetFormatFile, width, height, custPrpView, REMOVE_MODIFIED_NOTES) Then
        Err.Raise vbError, "", "Failed to set the sheet format"
    End If
    
End Sub

Sub ReplaceDraftingStandard(draw As SldWorks.DrawingDoc, sPath As String)
    Dim bRetVal         As Boolean
    Dim swModExt        As SldWorks.ModelDocExtension

    Set swModExt = draw.Extension
'    Debug.Print "Replaced overall drafting standard with " & sPath
    bRetVal = swModExt.LoadDraftingStandard(sPath)

End Sub

Sub SetAnnotationsToDefaultFont(draw As SldWorks.DrawingDoc)
    
    Dim sheet As SldWorks.sheet
    Dim swView As SldWorks.View
    Dim swNote As SldWorks.Note
    Dim swAnn As SldWorks.Annotation
    Dim swActiveView As SldWorks.View
    Dim modView As ModelView
    Dim bRet As Boolean
    Dim bRet2 As Boolean
    
    
    Set sheet = draw.GetCurrentSheet ' draw.getcurrentsheet
    Set swActiveView = draw.ActiveDrawingView ' draw.activedrawingview
    Set swView = draw.GetFirstView ' This is the drawing template - we change all annotations including these
    
    While Not swView Is Nothing
        ' get all the notes from this view
        Set swNote = swView.GetFirstNote
        draw.ClearSelection2 (True)
'        Debug.Print "File = " & swModel.GetPathName
        Do While Not swNote Is Nothing
            Set swAnn = swNote.GetAnnotation
            ' skip the title block annotations!
            If Not swAnn.OwnerType = 2 Then
                ' sets annotation to use the default font
                bRet = swAnn.SetTextFormat(0, True, Nothing)
                bRet = swAnn.Select2(True, 0)
            End If
            Set swNote = swNote.GetNext
        Loop
' set the next view
' Returns FALSE if trying to activate the drawing sheet
        bRet2 = draw.ActivateView(swView.GetName2):
        If False = bRet2 Then
            Debug.Assert sheet.GetName = swView.GetName2
            bRet2 = draw.ActivateSheet(swView.GetName2)
        End If
        Debug.Assert bRet2
        Set swView = swView.GetNextView
    Wend
    
End Sub

Sub DeleteAndAddLayers(stdLayerMap As Variant, swModel As SldWorks.DrawingDoc)
    Dim swLayerMgr                  As SldWorks.LayerMgr
    Dim vLayerArr                   As Variant
    Dim vLayer                      As Variant
    Dim swLayer                     As SldWorks.Layer
    Dim i                           As Integer

    Set swLayerMgr = swModel.GetLayerManager

    ' get current layers
    vLayerArr = swLayerMgr.GetLayerList
    ' delete all current layers
    For Each vLayer In vLayerArr
        Set swLayer = swLayerMgr.GetLayer(vLayer)
'        Debug.Print "    Layer          = " & swLayer.Name
'        Debug.Print "    Color          = " & swLayer.Color
'        Debug.Print "    Description    = " & swLayer.Description
'        Debug.Print "    ID             = " & swLayer.GetID
'        Debug.Print "    Style          = " & swLayer.Style
'        Debug.Print "    Visible        = " & swLayer.Visible
'        Debug.Print "    Width          = " & swLayer.Width
'        Debug.Print "    Printable      = " & swLayer.Printable
        swLayerMgr.DeleteLayer (vLayer)
    Next

    'add all layers
    For i = LBound(stdLayerMap) To UBound(stdLayerMap)
        Dim layerName As String
        Dim layerDesc As String
        Dim layerColor As Long
        Dim layerStyle As Long
        Dim layerWidth As Long
        Dim layerVal As Integer
        
        layerName = stdLayerMap(i)(0)
        layerDesc = stdLayerMap(i)(1)
        layerColor = stdLayerMap(i)(2)
        layerStyle = stdLayerMap(i)(3)
        layerWidth = stdLayerMap(i)(4)
'        Debug.Print "    Name           = " & stdLayerMap(i)(0)
'        Debug.Print "    Description    = " & stdLayerMap(i)(1)
'        Debug.Print "    Color          = " & stdLayerMap(i)(2)
'        Debug.Print "    Style          = " & stdLayerMap(i)(3)
'        Debug.Print "    Width          = " & stdLayerMap(i)(4)
        layerVal = swLayerMgr.AddLayer(layerName, layerDesc, layerColor, layerStyle, layerWidth)
    Next i
End Sub

Sub StartNoScreenUpdate(swModel As SldWorks.ModelDoc2)
    Dim modView As ModelView
    Set modView = swModel.ActiveView
    modView.EnableGraphicsUpdate = False
End Sub

Sub EndNoScreenUpdate(swModel As SldWorks.ModelDoc2)
    Dim modView As ModelView
    Set modView = swModel.ActiveView
    modView.EnableGraphicsUpdate = True
End Sub

Sub StartLockFMT(swModel As SldWorks.ModelDoc2)
    swModel.FeatureManager.EnableFeatureTree = False
End Sub

Sub EndLockFMT(swModel As SldWorks.ModelDoc2)
    swModel.FeatureManager.EnableFeatureTree = True
End Sub

Sub ReloadSheetFormat(sheet As SldWorks.sheet)
    Dim reloadResult As swReloadTemplateResult_e
    reloadResult = sheet.ReloadTemplate(False)
    Debug.Print "Reload sheet format for <" & sheet.GetName & ">: " & GetReloadResult(reloadResult)
End Sub

Private Function GetReloadResult(ByVal result As swReloadTemplateResult_e) As String
    Select Case result
    Case swReloadTemplate_Success
        GetReloadResult = "Success"
    Case swReloadTemplate_UnknownError
        GetReloadResult = "FAIL - Unknown Error"
    Case swReloadTemplate_FileNotFound
        GetReloadResult = "FAIL - File Not Found"
    Case swReloadTemplate_CustomSheet
        GetReloadResult = "FAIL - Custom Sheet"
    Case swReloadTemplate_ViewOnly
        GetReloadResult = "FAIL - View Only"
    Case Else
        GetReloadResult = "FAIL - <unrecognized error code - " & result & ">"
    End Select
End Function
