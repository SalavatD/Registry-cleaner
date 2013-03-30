Enumeration
  #WINDOW
EndEnumeration

Enumeration
  #MENU_BAR
EndEnumeration

Enumeration
  #MENU_EXIT
  #MENU_ABOUT
EndEnumeration

Enumeration
  #TEXT_SEARCH_KEYWORDS
  #STRING_SEARCH_KEYWORDS
  #BUTTON_START_STOP_SEARCH
  #LIST_ICON
  #BUTTON_CHECK_ALL
  #BUTTON_UNCHECK_ALL
  #BUTTON_DELETE
  #CHECK_BOX_HKLM
  #CHECK_BOX_HKCU
  #CHECK_BOX_HKCC
  #CHECK_BOX_HKU
  #CHECK_BOX_HKCR
EndEnumeration

Enumeration
  #STATUS_BAR
EndEnumeration

Enumeration
  #APPROXIMATE_NUMBER_OF_KEYS_IN_HKCR = 150000
  #APPROXIMATE_NUMBER_OF_KEYS_IN_HKCU = 6000
  #APPROXIMATE_NUMBER_OF_KEYS_IN_HKU = 12000
  #APPROXIMATE_NUMBER_OF_KEYS_IN_HKCC = 100
  #APPROXIMATE_NUMBER_OF_KEYS_IN_HKLM = 300000
EndEnumeration

Enumeration
  #PROGRESS_DELAY = 250
EndEnumeration

Global hkcrSearchThread
Global hkcuSearchThread
Global hkuSearchThread
Global hkccSearchThread
Global hklmSearchThread

Global countOfFound
Global countOfScanned
Global progressBarTotal

Global Dim searchKeywords.s(0)

inSearchingProcess = #False
searchString.s = ""

Procedure KillAllSearchThreads()
  If IsThread(hkcrSearchThread)
    KillThread(hkcrSearchThread)
  EndIf 
  If IsThread(hkcuSearchThread)
    KillThread(hkcuSearchThread)
  EndIf 
  If IsThread(hkuSearchThread)
    KillThread(hkuSearchThread)
  EndIf 
  If IsThread(hkccSearchThread)
    KillThread(hkccSearchThread)
  EndIf 
  If IsThread(hklmSearchThread)
    KillThread(hklmSearchThread)
  EndIf
EndProcedure

Procedure TokenizeString(Array searchKeywords$(1), searchString$, delimeter$)
  countOfKeywords = CountString(searchString$, delimeter$) + 1
  Dim searchKeywords$(countOfKeywords)
  For i = 1 To countOfKeywords
    searchKeywords$(i - 1) = StringField(searchString$, i, delimeter$)
  Next
  ProcedureReturn countOfKeywords
EndProcedure

Procedure SearchProcedure(*Point)
  startKey.s = PeekS(*Point)
  valueCounter = 0
  While RegListSubValue(startKey, valueCounter, ".") <> ""
    subValue.s = RegListSubValue(startKey, valueCounter, ".")
    For i = 0 To ArraySize(searchKeywords())
      If FindString(LCase(subValue), LCase(searchKeywords(i)), 0)
        AddGadgetItem(#LIST_ICON, -1, startKey +"|"+ subValue)
      EndIf
    Next 
    regValue.s = RegGetValue(startKey, subValue, ".")
    For i = 0 To ArraySize(searchKeywords())
      If FindString(LCase(regValue), LCase(searchKeywords(i)), 0)
        AddGadgetItem(#LIST_ICON, -1, startKey + "|" + subValue + "=" + regValue)
      EndIf
    Next
    valueCounter + 1
    countOfScanned + 1
  Wend
  keyCounter = 0
  While RegListSubKey(startKey, keyCounter, ".") <> ""
    subKey.s = RegListSubKey(startKey, keyCounter, ".")
    For i = 0 To ArraySize(searchKeywords())
      If FindString(LCase(subKey), LCase(searchKeywords(i)), 0)
        AddGadgetItem(#LIST_ICON, -1, startKey + subKey + "\") 
      EndIf
    Next 
    startSububKey.s = startKey + subKey + "\"
    SearchProcedure(@startSububKey)
    keyCounter + 1
    countOfScanned + 1
  Wend
EndProcedure

Procedure CreateSearchThreads()
  If GetGadgetState(#CHECK_BOX_HKCR)
    hkcrSearchThread = CreateThread(@SearchProcedure(), @"HKEY_CLASSES_ROOT\")
    progressBarTotal + #APPROXIMATE_NUMBER_OF_KEYS_IN_HKCR
  EndIf
  If GetGadgetState(#CHECK_BOX_HKCU)
    hkcuSearchThread = CreateThread(@SearchProcedure(), @"HKEY_CURRENT_USER\")
    progressBarTotal + #APPROXIMATE_NUMBER_OF_KEYS_IN_HKCU
  EndIf
  If GetGadgetState(#CHECK_BOX_HKU)
    hkuSearchThread = CreateThread(@SearchProcedure(), @"HKEY_USERS\")
    progressBarTotal + #APPROXIMATE_NUMBER_OF_KEYS_IN_HKU
  EndIf
  If GetGadgetState(#CHECK_BOX_HKCC)
    hkccSearchThread = CreateThread(@SearchProcedure(), @"HKEY_CURRENT_CONFIG\")
    progressBarTotal + #APPROXIMATE_NUMBER_OF_KEYS_IN_HKCC
  EndIf
  If GetGadgetState(#CHECK_BOX_HKLM)
    hklmSearchThread = CreateThread(@SearchProcedure(), @"HKEY_LOCAL_MACHINE\")
    progressBarTotal + #APPROXIMATE_NUMBER_OF_KEYS_IN_HKLM
  EndIf
EndProcedure

Procedure UpdateProgress(null)
  SetGadgetText(#BUTTON_START_STOP_SEARCH, "Стоп")
  Repeat
    StatusBarText(#STATUS_BAR, 0, "Просканировано: " + Str(countOfScanned) + " ключей") 
    StatusBarProgress(#STATUS_BAR, 1, countOfScanned, #PB_Ignore, 0, progressBarTotal)
    If IsThread(hkcrSearchThread) = 0 And IsThread(hkcuSearchThread) = 0 And IsThread(hkuSearchThread) = 0 And IsThread(hkccSearchThread) = 0 And IsThread(hklmSearchThread) = 0
      StatusBarText(#STATUS_BAR, 0, "Поиск завершен. Просканировано: " + Str(countOfScanned) + " ключей")
      StatusBarProgress(#STATUS_BAR, 1, 0)
      SetGadgetText(#BUTTON_START_STOP_SEARCH, "Поиск")
      Break
    EndIf
    Delay(#PROGRESS_DELAY)
  ForEver
EndProcedure

Procedure OpenInRegEdit()
  selectedItem = GetGadgetState(#LIST_ICON)
  If selectedItem <> -1
    selectedItemPath$ = GetGadgetItemText(#LIST_ICON, selectedItem)
    If Right(selectedItemPath$, 1) = "\"
      selectedItemPath$ = RTrim(selectedItemPath$, "\")
    Else
      strstr_(String$, StringToFind$)
      If FindString(selectedItemPath$, "\|", 0)
        selectedItemPath$ = Mid(selectedItemPath$, 0, FindString(selectedItemPath$,"\|", 0) - 1)
      EndIf
    EndIf
    RegCreateKeyValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", "LastKey", selectedItemPath$, #REG_SZ, ".")
    RunProgram("regedit")
  EndIf
EndProcedure

Procedure CheckAll()
  If CountGadgetItems(#LIST_ICON)
    For i = 0 To CountGadgetItems(#LIST_ICON)
      SetGadgetItemState(#LIST_ICON, i, #PB_ListIcon_Checked)
    Next
  EndIf
EndProcedure

Procedure UncheckAll()
  If CountGadgetItems(#LIST_ICON)
    For i = 0 To CountGadgetItems(#LIST_ICON)
      SetGadgetItemState(#LIST_ICON, i, ~#PB_ListIcon_Checked)
    Next
  EndIf
EndProcedure

Procedure DeleteCheckedItems()
  deletedItemsCounter = 0
  For i = 0 To CountGadgetItems(#LIST_ICON)
    If GetGadgetItemState(#LIST_ICON, i) & #PB_ListIcon_Checked
      deletedItemsCounter + 1
      checkedItemPath$ = GetGadgetItemText(#LIST_ICON, i) 
      StatusBarText(#STATUS_BAR, 0, "Удаление:"+ checkedItemPath$)
      If Right(checkedItemPath$, 1) = "\"
        checkedItemPath$ = RTrim(checkedItemPath$, "\")
        RegDeleteKeyWithAllSub(checkedItemPath$,".")
        RemoveGadgetItem(#LIST_ICON, i)
        i - 1
      Else
        If FindString(checkedItemPath$,"\|", 0)
          position = FindString(checkedItemPath$,"\|", 0)
          value$ = Mid(checkedItemPath$, 0, position - 1)
          key$ = Mid(checkedItemPath$, position + 2)
          If FindString(key$, "=", 0)
            key$ = Mid(key$, 0, FindString(key$, "=", 0) - 1)
          EndIf
          RegDeleteValue(value$, key$, ".") 
          RemoveGadgetItem(#LIST_ICON, i)
          i - 1
        EndIf
      EndIf
    EndIf
  Next
  StatusBarText(#STATUS_BAR, 0, "Удаление завершено")
EndProcedure

Procedure OpenMainWindow()
  If OpenWindow(#WINDOW, #PB_Ignore, #PB_Ignore, 800, 600, "Очистка реестра", #PB_Window_MinimizeGadget | #PB_Window_ScreenCentered)
    If CreateMenu(#MENU_BAR, WindowID(#WINDOW))
      MenuTitle("Файл")
      MenuItem(#MENU_EXIT, "Выход")
      MenuTitle("Справка")
      MenuItem(#MENU_ABOUT, "О программе")
    EndIf
    If CreateStatusBar(#STATUS_BAR, WindowID(#WINDOW))
      AddStatusBarField(600)
      AddStatusBarField(200)
      StatusBarText(#STATUS_BAR, 0, "Нажмите 'Поиск' для начала сканирования")
      StatusBarProgress(#STATUS_BAR, 1, 0)
    EndIf
    TextGadget(#TEXT_SEARCH_KEYWORDS, 5, 10, 190, 20, "Ключевые слова (через запятую):", #PB_Text_Right)
    StringGadget(#STRING_SEARCH_KEYWORDS, 200, 5, 400, 25, "")
    ButtonGadget(#BUTTON_START_STOP_SEARCH, 605, 5, 190, 25, "Поиск")
    ListIconGadget(#LIST_ICON, 5, 35, 790, 475, "Ключ реестра", 750, #PB_ListIcon_CheckBoxes | #PB_ListIcon_GridLines)
    ButtonGadget(#BUTTON_CHECK_ALL, 5, 515, 100, 35, "Отметить все")
    ButtonGadget(#BUTTON_UNCHECK_ALL, 110, 515, 100, 35, "Снять все")
    CheckBoxGadget(#CHECK_BOX_HKLM, 215, 515, 90, 35, "HKLM")
    GadgetToolTip(#CHECK_BOX_HKLM, "HKEY_LOCAL_MACHINE")
    SetGadgetState(#CHECK_BOX_HKLM, 1)
    CheckBoxGadget(#CHECK_BOX_HKCU, 310, 515, 90, 35, "HKCU")
    GadgetToolTip(#CHECK_BOX_HKCU, "HKEY_CURRENT_USER")
    SetGadgetState(#CHECK_BOX_HKCU, 1)
    CheckBoxGadget(#CHECK_BOX_HKCC, 405, 515, 90, 35, "HKCC")
    GadgetToolTip(#CHECK_BOX_HKCC, "HKEY_CURRENT_CONFIG")
    SetGadgetState(#CHECK_BOX_HKCC, 1)
    CheckBoxGadget(#CHECK_BOX_HKU, 500, 515, 90, 35, "HKU")
    GadgetToolTip(#CHECK_BOX_HKU, "HKEY_USERS")
    SetGadgetState(#CHECK_BOX_HKU, 1)
    CheckBoxGadget(#CHECK_BOX_HKCR, 595, 515, 90, 35, "HKCR")
    GadgetToolTip(#CHECK_BOX_HKCR, "HKEY_CLASSES_ROOT")
    ButtonGadget(#BUTTON_DELETE, 690, 515, 105, 35, "Удалить")
  EndIf
EndProcedure

OpenMainWindow()

Repeat
  event = WaitWindowEvent()
  eventMenu = EventMenu()
  eventGadget = EventGadget()
  eventType = EventType()
  
  If event = #PB_Event_Menu
    If eventMenu = #MENU_EXIT
      Break
    ElseIf eventMenu = #MENU_ABOUT
      MessageRequester("О программе", "Очистка реестра. Версия 1.0" + #CR$ + #CR$ + "Автор: Салават Даутов" + #CR$ + #CR$ + "Дата создания: март 2013", #MB_ICONINFORMATION)
    EndIf
  EndIf
  
  If event = #PB_Event_Gadget
    If eventGadget = #BUTTON_START_STOP_SEARCH
      If inSearchingProcess
        inSearchingProcess = #False
        KillAllSearchThreads()
      Else
        inSearchingProcess = #True
        ClearGadgetItems(#LIST_ICON)
        countOfScanned = 0
        progressBarTotal = 0
        searchString = GetGadgetText(#STRING_SEARCH_KEYWORDS)
        TokenizeString(searchKeywords(), searchString, ",")
        CreateSearchThreads()
        CreateThread(@UpdateProgress(), 0)
      EndIf
    ElseIf eventGadget = #LIST_ICON
      If eventType = #PB_EventType_LeftDoubleClick
        OpenInRegEdit()
      EndIf
    ElseIf eventGadget = #BUTTON_CHECK_ALL
      CheckAll()
    ElseIf eventGadget = #BUTTON_UNCHECK_ALL
      UncheckAll()
    ElseIf eventGadget = #BUTTON_DELETE
      DeleteCheckedItems()
    EndIf
  EndIf
Until event = #PB_Event_CloseWindow

; IDE Options = PureBasic 4.51 (Windows - x86)
; CursorPosition = 302
; FirstLine = 261
; Folding = --
; EnableThread
; EnableXP
; EnableAdmin
; UseIcon = Icon.ico
; Executable = Очистка реестра.exe
; SubSystem = UserLibThreadSafe
; DisableDebugger
; IncludeVersionInfo
; VersionField0 = 1.0.0.0
; VersionField1 = 1.0.0.0
; VersionField2 = Салават Даутов
; VersionField3 = Очистка реестра
; VersionField4 = 1.0
; VersionField5 = 1.0
; VersionField6 = Очистка реестра
; VersionField7 = Очистка реестра
; VersionField8 = Очистка реестра.exe
; VersionField17 = 0419 Russian
