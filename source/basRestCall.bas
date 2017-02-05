Attribute VB_Name = "basRde"
'------------------------------------------------------------------------
' Description  : getting results vom similarity calculation
'------------------------------------------------------------------------
'
'Declarations

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : requests similarity analysis from RDE
' Parameter     : pstrInputData     - headline + 2 line of product data
' Returnvalue   : results of the similarity comparison
'-------------------------------------------------------------
Public Function getSimilarityComparison(pstrInputData As String) As String

    Dim objServerXml As New MSXML2.ServerXMLHTTP
    Dim strSimRequest As String
                     
    On Error GoTo error_handler
    basSystem.log "requesting RDE for similarity analysis", cLogInfo
    'TODO: put config at least to basConstants
    strSimRequest = "http://rde-dev-1602.lhotse.ov.otto.de:8080/rde_server/admin/res/Dings2/datacleansing/similarity/tasks/sim1/analyze"
    basSystem.log "input data:" & vbCrLf & pstrInputData, cLogDebug
    With objServerXml
        .Open "POST", strSimRequest, False, "admin", "admin"
        .setRequestHeader "Accept", "text/csv"
        .setRequestHeader "Content-Type", "text/csv"
        .send pstrInputData
        .waitForResponse 60
    End With
    If objServerXml.Status = 200 Then
        basSystem.log "got analysis result from RDE", cLogInfo
        getSimilarityComparison = objServerXml.responseText
    Else
        basSystem.log_error "basRde.getSimilarityComparison", "RDE failed to answer request with return code " & objServerXml.Status
        getSimilarityComparison = ""
    End If
    Exit Function
                     
error_handler:
    basSystem.log_error "basRde.getSimilarityComparison"
End Function
'-------------------------------------------------------------
' Description   : converts a comparison result into a collection
' Parameter     : pstrSimResult     - result of a similarity comparison
' Returnvalue   : colection containing the distinct values of the result
'-------------------------------------------------------------
Public Function getCollectionFromSimResult(pstrSimResult As String) As Collection

    Dim arrResultData As Variant
    Dim arrResultHeader As Variant
    Dim strResultData As String
    Dim strResultHeader As String
    Dim colResultData As New Collection
    Dim intItem As Integer
                     
    On Error GoTo error_handler
    strResultHeader = Left(pstrSimResult, InStr(1, pstrSimResult, vbLf) - 2)
    basSystem.log "Header:" & vbCrLf & ">" & strResultHeader & "<", cLogDebug
    arrResultHeader = Split(strResultHeader, "|")
    strResultData = Right(pstrSimResult, Len(pstrSimResult) - InStr(1, pstrSimResult, vbLf) + 2)
    strResultData = Replace(strResultData, vbLf, "")
    strResultData = Replace(strResultData, vbCr, "")
    basSystem.log "Data:" & vbCrLf & ">" & strResultData & "<", cLogDebug
    arrResultData = Split(strResultData, "|")
    For intItem = 0 To UBound(arrResultHeader)
        colResultData.Add arrResultData(intItem), arrResultHeader(intItem)
    Next
    Set getCollectionFromSimResult = colResultData
    Exit Function
                     
error_handler:
    basSystem.log_error "basRde.getSimilarityComparison"
End Function
