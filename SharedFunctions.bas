Attribute VB_Name = "SharedFunctions"
Option Explicit
'//////////////////////////////////////////////////////////
'//
'// SurfTabs
'// A Tiger Studios product
'// By Jim Black
'// Jim@BlacksWeb.com
'//
'//////////////////////////////////////////////////////////

Global G_LoadFinished As Boolean
Global CurTab_Index As Integer
Global MaxTab_Index As Integer
Global AddURL, ClickHistory As Boolean
Global FreedBrowserArray() As Integer
Global FreedBrowserArrayEmpty As Boolean
Global BrowserArrayEmpty As Boolean
Global OpenOk As Boolean
Global OpenURL As String
Global OpenNewTab As Boolean
Global SetFocusOnly As Boolean
Global InTabSetFocus As Boolean
Global Deleting As Boolean
Global Found As Boolean
Global URLcount As Integer
Global strLocationName As String
Global strLocationURL As String
Global tempString As String
Global gProgPath As String
Global lLeft, lTop, lWidth, lHeight As Long
Global ViewingFavorites As Boolean
Global UnLoading As Boolean
Global HistoryFileChanged As Boolean
Global bBackSpace As Boolean
Global bDeleteKey As Boolean

'Options
Global OptionsSaved As Boolean
'General
Global gMinToSysTray As Long
Global gMultipleInstances As Long
Global gDefaultBrowser As Long
'Browser Tabs
Global gBrowserTitleLength As Long
Global gtxtBrowserTitleLength As String
Global gchkRefreshBrowser As Long
Global gtxtRefreshBrowser As String
'New Tabs
Global gNewTabHome As Long
Global gNewTabSearch As Long
Global gNewTabAddressTyped As Long
Global gNewTabFavorites As Long
Global gNewTabHistory As Long
Global gDefaultNewButton As Long
'Startup
Global gStartPage As Long


Function GetFormInfo()
    'GET THE POSITION AND SIZE OF BROWSER FORM
    If getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserTop") <> "" Then
        frmBrowser.Top = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserTop")
        frmBrowser.Left = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserLeft")
        frmBrowser.Width = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserWidth")
        frmBrowser.Height = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserHeight")
        G_LoadFinished = True
    End If

End Function

Function SaveFormInfo()
    'SAVE THE POSITION AND SIZE OF BROWSER FORM
    If frmBrowser.WindowState = 0 Then
        Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserTop", Str(frmBrowser.Top))
        Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserLeft", Str(frmBrowser.Left))
        Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserWidth", Str(frmBrowser.Width))
        Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Position", "BrowserHeight", Str(frmBrowser.Height))
    End If

End Function

Function GetSavedTabURLs()
On Error GoTo GetSavedTabURLs_Error:
    Dim X As Integer
    X = 0
    Dim URL As String
    
    URLcount = Int(getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs", "Count"))
    
    If URLcount > 0 Then
        'Set the current tabs browser to saved URL
        frmBrowser.TabStrip1.Tabs(CurTab_Index).Caption = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs", "1Location")
        frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).Visible = True
        URL = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs", Str(X + 1) + "URL")
        frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).Navigate URL
        frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).ZOrder (1)
        X = 1
        
        While X <= URLcount - 1
            'Increment WebBrowser indexes
            MaxTab_Index = MaxTab_Index + 1
            CurTab_Index = CurTab_Index + 1
            frmBrowser.TabStrip1.Tabs.Add
            frmBrowser.TabStrip1.Tabs.Item(CurTab_Index).Caption = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs", Str(X + 1) + "Location")
            frmBrowser.TabStrip1.Tabs.Item(CurTab_Index).Selected = True
            frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag = CurTab_Index - 1
            'Load new browser, and enable it
            Load frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag)
            frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).Visible = True
            frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).ZOrder (1)
            URL = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs", Str(X + 1) + "URL")
            frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).Navigate URL
            X = X + 1
        Wend
        
        MaxTab_Index = URLcount
        Call DeleteKey(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs")
    
    End If
    Exit Function

GetSavedTabURLs_Error:
    ShowErrorMessageBox ("GetSavedTabURLs")
End Function

Function SaveCurrentTabURLs()
On Error GoTo SaveCurrentTabURLs_Error:
    Dim X As Integer
    
    If frmBrowser.WindowState = 0 Then
        
        For X = 1 To frmBrowser.TabStrip1.Tabs.Count
            Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs", Str(X) + "URL", frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(X).Tag).LocationURL)
            Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs", Str(X) + "Location", frmBrowser.TabStrip1.Tabs(X).Caption)
        Next
        
        Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\SavedURLs", "Count", Str(X - 1))
    
    End If
    Exit Function

SaveCurrentTabURLs_Error:
    ShowErrorMessageBox ("SaveCurrentTabURLs")
End Function

Public Sub PrintBrowser()
    frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).SetFocus
    SendKeys "^p"

End Sub

Sub AddToFavorites()
    Dim shellHelper As New ShellUIHelper
    Dim strLocationName, strLocationURL As String
    strLocationName = frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).LocationName
    strLocationURL = frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(CurTab_Index).Tag).LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End Sub

Sub SaveOptions()
On Error GoTo SaveOptions_Error:
    '/////////////////////////////////////////////////////////////////
    'General tab
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkMinToSysTray", frmOptions.chkMinToSysTray.Value)
    gMinToSysTray = frmOptions.chkMinToSysTray.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkMultipleInstances", frmOptions.chkMultipleInstances.Value)
    gMultipleInstances = frmOptions.chkMultipleInstances.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkDefaultBrowser", frmOptions.chkDefaultBrowser.Value)
    gDefaultBrowser = frmOptions.chkDefaultBrowser.Value
        
    '/////////////////////////////////////////////////////////////////
    'Browser Tabs tab
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkBrowserTitleLength", frmOptions.chkBrowserTitleLength.Value)
    gBrowserTitleLength = frmOptions.chkBrowserTitleLength.Value
    
    Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "txtBrowserTitleLength", frmOptions.txtBrowserTitleLength.text)
    gtxtBrowserTitleLength = frmOptions.txtBrowserTitleLength.text
    gchkRefreshBrowser = frmOptions.chkRefreshBrowser.Value
    gtxtRefreshBrowser = frmOptions.txtRefreshBrowser.text
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkRefreshBrowser", gchkRefreshBrowser)
    Call savestring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "txtRefreshBrowser", gtxtRefreshBrowser)
    
    '/////////////////////////////////////////////////////////////////
    'New Tabs tab
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabHome", frmOptions.chkNewTabHome.Value)
    gNewTabHome = frmOptions.chkNewTabHome.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabSearch", frmOptions.chkNewTabSearch.Value)
    gNewTabSearch = frmOptions.chkNewTabSearch.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabAddressTyped", frmOptions.chkNewTabAddressTyped.Value)
    gNewTabAddressTyped = frmOptions.chkNewTabAddressTyped.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabFavorites", frmOptions.chkNewTabFavorites.Value)
    gNewTabFavorites = frmOptions.chkNewTabFavorites.Value
    
    Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabHistory", frmOptions.chkNewTabHistory.Value)
    gNewTabHistory = frmOptions.chkNewTabHistory.Value
    
    If frmOptions.optDefaultNewButton(0).Value = True Then
        Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "optDefaultNewButton", 0)
        gDefaultNewButton = 0
    Else
        If frmOptions.optDefaultNewButton(1).Value = True Then
            Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "optDefaultNewButton", 1)
            gDefaultNewButton = 1
        Else
            Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "optDefaultNewButton", 2)
            gDefaultNewButton = 2
        End If
    End If
        
    '/////////////////////////////////////////////////////////////////
    'Startup tab
    If frmOptions.optStartPage(0).Value = True Then
        Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "optStartPage", 0)
        gStartPage = 0
    Else
        If frmOptions.optStartPage(1).Value = True Then
            Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "optStartPage", 1)
            gStartPage = 1
        Else
            Call SaveDword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "optStartPage", 2)
            gStartPage = 2
        End If
    End If
    Exit Sub

SaveOptions_Error:
    ShowErrorMessageBox ("SaveOptions")
End Sub

Sub GetOptions()
On Error GoTo GetOptions_Error:

    '/////////////////////////////////////////////////////////////////
    'General tab
    gMinToSysTray = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkMinToSysTray")
    gMultipleInstances = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkMultipleInstances")
    gDefaultBrowser = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkDefaultBrowser")
    
    '/////////////////////////////////////////////////////////////////
    'Browser Tabs tab
    gBrowserTitleLength = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkBrowserTitleLength")
    gtxtBrowserTitleLength = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "txtBrowserTitleLength")
    gchkRefreshBrowser = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkRefreshBrowser")
    gtxtRefreshBrowser = getstring(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "txtRefreshBrowser")
    
    '/////////////////////////////////////////////////////////////////
    'New Tabs tab
    gNewTabHome = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabHome")
    gNewTabSearch = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabSearch")
    gNewTabAddressTyped = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabAddressTyped")
    gNewTabFavorites = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabFavorites")
    gNewTabHistory = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "chkNewTabHistory")
    gDefaultNewButton = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "optDefaultNewButton")
    
    '/////////////////////////////////////////////////////////////////
    'Startup tab
    gStartPage = getdword(HKEY_CURRENT_USER, "Software\Tiger Studios\SurfTabs\Options", "optStartPage")
    Exit Sub

GetOptions_Error:
    ShowErrorMessageBox ("GetOptions")
End Sub
Sub SetProgPath()
On Error GoTo SetProgPath_Error:
    ' If dragged file is in the root, append filename.
    If Mid(App.Path, Len(App.Path)) = "\" Then
        gProgPath = App.Path
    ' If dragged file is not in root, append "\" and filename.
    Else
        gProgPath = App.Path & "\"
    End If
    Exit Sub

SetProgPath_Error:
    ShowErrorMessageBox ("SetProgPath")
End Sub

Sub Get_History()
On Error GoTo error
    'Reads in history file and populates address list
    Open gProgPath & "History.dat" For Input As #1
    frmBrowser.cboAddress.Clear
    
    Do Until EOF(1)
        Line Input #1, tempString
        frmBrowser.cboAddress.AddItem tempString
    Loop
    
    Close #1
    Exit Sub

error:
    MsgBox "Error loading history file." & vbCrLf & _
            "Please make sure it exists (even if it is blank)." & vbCrLf & _
            "The history file should be " & gProgPath & "History.dat.", vbCritical
End Sub

Sub Save_History()
On Error GoTo Save_History_Error:
'/////////////////////////////////////////////////////////////
'//
'// NOTE: THIS IS NO LONGER BEING USED
'// SAVING HISTORY TO REGISTRY, SAME AS IE
'// IN THE TypedURLs key
'//
'////////////////////////////////////////////////////////////
    'SAVE IT!
    Dim X As Integer
    Open gProgPath & "History.dat" For Output As #1
        
    For X = 0 To frmBrowser.cboAddress.ListCount - 1
        frmBrowser.cboAddress.ListIndex = X
        Print #1, frmBrowser.cboAddress.text
    Next
    
    Close #1
    Exit Sub

Save_History_Error:
    MsgBox "An error occured while saving the history file", vbCritical
End Sub

Sub RepositionProgressBar()
        'RESIZE and POSITION THE PROGRESS BAR
        lLeft = frmBrowser.StatusBar.Panels(2).Left + 10
        lTop = frmBrowser.StatusBar.Top + 30
        lWidth = frmBrowser.StatusBar.Panels(2).Width - 20
        lHeight = frmBrowser.StatusBar.Height - 40
        frmBrowser.ProgressBar.Move lLeft, lTop, lWidth, lHeight

End Sub

Sub GetTypedURLs()
On Error GoTo GetTypedURLs_Error:
    'Get the TypedURLs from the registry
    'And populate the address list
    Dim Done As Boolean
    Dim URLnum As Integer
    
    Done = False
    URLnum = -1
    
    While Not Done
        
        URLnum = URLnum + 1
        tempString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", "url" & URLnum)
        
        If URLnum = 0 And tempString = "" Then
            'Do nothing
        Else
            
            If tempString <> "" Then
                frmBrowser.cboAddress.AddItem tempString
            Else
                Done = True
            End If
        
        End If
    
    Wend
    Exit Sub

GetTypedURLs_Error:
    ShowErrorMessageBox ("GetTypedURLs")
End Sub

Sub SaveTypedURLs()
On Error GoTo SaveTypedURLs_Error:
    'Save the TypedURLs in the address list
    'To the registry in the TypedURLs key used by IE
    Dim Done As Boolean
    Dim URLnum As Integer
    
    For URLnum = 0 To frmBrowser.cboAddress.ListCount - 1
        frmBrowser.cboAddress.ListIndex = URLnum
        tempString = frmBrowser.cboAddress.text
        Call savestring(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", "url" & URLnum, tempString)
    Next
    Exit Sub

SaveTypedURLs_Error:
    ShowErrorMessageBox ("SaveTypedURLs")
End Sub

Sub MoveBrowserOffFormTab(index As Integer)
    frmBrowser.brwWebBrowser(index).Left = frmBrowser.ScaleLeft + 10000000
    
End Sub

Sub SelectBrowserTab(index As Integer)
    frmBrowser.TabStrip1.Tabs(index).Selected = True
    frmBrowser.TabStrip1.SetFocus
    frmBrowser.TabStrip1.Refresh
    frmBrowser.brwWebBrowser(frmBrowser.TabStrip1.Tabs(index).Tag).SetFocus
'jrb may need to do form resize here

End Sub

Sub ShowErrorMessageBox(Where As String)
    'MsgBox ("Error in " & Where & ", " & vbCrLf & _
            "Please report the bug to SurfTabs web site http://BlacksWeb.com/SurfTabs" & vbCrLf & _
            "Or contact Jim@BlacksWeb.com.  Thank you.")

End Sub
