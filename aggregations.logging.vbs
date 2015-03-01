Option Explicit


'***********************************************************
' aggregating q:\utils\libs\extensionsbase.vbs

'###############################################################################
' Library: ExtensionsBase.vbs
'
' About: 
'  Basic extensions to VBScript/QTP 
'  Things that *should* have been in VBScript but aren't
'  (Except for strings, files and dates... this file got too big so they got moved)
'
'  Copyright (C) 2008, 2009, 2010 Akien MacIain
'
'  This program is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' Function/Dependencies: ReadPrefsFileIntoDict/class-lists.vbs
'
'###############################################################################
'Option Explicit

'===============================================================================
' Section: Public Functions
'   Functions that are exported by the library
'===============================================================================

'-------------------------------------------------------------------------------
' Function: FrameworkDetectExtensionsBase
'   Utility function for the Framework Compilation checking utility
'
' Returns:
'  (integer) always returns 1
'-------------------------------------------------------------------------------
Function FrameworkDetectExtensionsBase()
	FrameworkDetectExtensionsBase = 1
End Function


'============================================================================================================
' MISCELLANIOUS



'-------------------------------------------------------------------------------
' Function: MakeHotTimer
'
'  Returns a running timer object
'
' Parameters:
'   - TimerName - The name of the timer to create
'
' Returns:
'  - object - A MercuryTimer object which has been started (is "hot", which is to say, running)
'
' NOTE:
'  In QTP a MercuryTimer is an object who's name is a string. This function adds the 
'  start time to the string.
'-------------------------------------------------------------------------------
'<<<-------TIMER GLOBAL SHARED DATA TO INSURE TIMER UNIQUENESS------->>>
'This module global timer assures we get unique timer names for each instance,
'even in the case of threaded code that overlaps calls
On Error Resume Next
Private vbscript_extensions_timerSinceStartup
Set vbscript_extensions_timerSinceStartup = MercuryTimers.Timer("ScriptExecutionDurationTimer")
vbscript_extensions_timerSinceStartup.Reset
vbscript_extensions_timerSinceStartup.Start
On Error Goto 0

Public Function MakeHotTimer(timerName)
   Dim MyTimer, stringName
   stringName = timerName & " - " & vbscript_extensions_timerSinceStartup.ElapsedTime
   Set MyTimer = MercuryTimers.Timer(stringName)
   'LogTrace "routine=MakeHotTimer;message=Creating timer: " & stringName
   On Error Resume Next
   MyTimer.Start
   MyTimer.Stop
   MyTimer.Reset
   On Error Goto 0
   MyTimer.Start
   Set MakeHotTimer = MyTimer
End Function

'-------------------------------------------------------------------------------
' Function: WSSHell
'  
'  Returns a WSSHell object. Uses the GlobalDictionary to store it so it 
'  doesn't have to recreate it every time
'
' Parameters:
'   - None
'
' Returns:
'  The WSSHell
'
' Notes:
'  Assumes a GlobalDictionary to store it in
'
'-------------------------------------------------------------------------------
Function WSSHell()
	If IsEmpty(GlobalDictionary("wsshell")) Then
		GlobalDictionaryAdd "wsshell",CreateObject("WScript.Shell")
	End If
	Set WSSHell = GlobalDictionary("wsshell")
End Function

'============================================================================================================
' ASSIGN GROUP

'-------------------------------------------------------------------------------
' Sub: Assign
'
'  Shortcut that performs an assignment after determining if the value to 
'  set is an object (and so using the Set keyword)
'
' Parameters:
'   - variableToSet - the item to assign the value to
'   - valueToSet  - the value to assign to the item
'
' Usage:
'  Assign foo, bar
'
'-------------------------------------------------------------------------------
Sub Assign (byref variableToSet, valueToSet)
   If IsObject(valueToSet) Then
      Set variableToSet = valueToSet
   Else
      variableToSet = valueToSet
   End If
End Sub

'-------------------------------------------------------------------------------
' Sub: AssignByPriority
'
'  Shortcut that performs an assignment based on several options, some of which
'  might be empty or NULL. Uses the first value in the list which is not empty
'  or NULL.
'
' Parameters:
'   - variableToSet - variable to assign to
'   - vopt1         - option 1
'   - vopt2         - option 2
'   - vopt3         - option 3
'   - vopt4         - option 4
'
' Usage:
'  AssignByPriority phoneNumberToUse, cellNumber, homeNumber, workNumber, auntMarthasNumber
'
'-------------------------------------------------------------------------------
Sub AssignByPriority (byref variableToSet, ByVal opt1, ByVal opt2, ByVal opt3, ByVal opt4)
	AssignIfNotEmpty variableToSet, opt4
	AssignIfNotEmpty variableToSet, opt3
	AssignIfNotEmpty variableToSet, opt2
	AssignIfNotEmpty variableToSet, opt1
End Sub

'-------------------------------------------------------------------------------
' Sub: AssignIfNotEmpty
'
'  Shortcut that performs an assignment after determining if the valueToSet
'  variable is set to something OTHER THAN Empty, "", or NULL. If valueToSet
'  is set to any of those values, variableToSet will be unchanged on exit
'
' Parameters:
'   - variableToSet - the item to assign the value to
'   - valueToSet  - the value to assign to the item
'
' Usage:
'  AssignIfNotEmpty foo, bar
'
'-------------------------------------------------------------------------------
Sub AssignIfNotEmpty (byref variableToSet, ByVal valueToSet) ' Assigns IF THE *valueToSet* IS NOT EMPTY
	If (IsEmpty(valueToSet) OR (valueToSet = "") OR IsNull(valueToSet)) = False Then
		variableToSet = valueToSet
	End If
End Sub

'-------------------------------------------------------------------------------
' Sub: AssignIfNull 
'
'  Part of the Assign family, AssignIfNull assigns "valueToSet" to the passed "variableToSet"
'  if "variableToSet" is NULL on entry 
'
' Parameters:
'   - variableToSet    - the item to assign the value to
'   - valueToSet  - the value to assign to the item
'
' Usage:
'  AssignIfNull foo, bar
'
'-------------------------------------------------------------------------------
Sub AssignIfNull(ByRef variableToSet, valueToSet)
	If IsNull(variableToSet) Then
		Assign variableToSet, valueToSet
	End If
End Sub

'-------------------------------------------------------------------------------
' Sub: AssignIfEmpty 
'
'  Part of the Assign family, AssignIfEmpty assigns "variableToSet" to the passed "valueToSet"
'  if "variableToSet" is NULL or Empty or "" on entry 
'
' Parameters:
'   - variableToSet    - the item to assign the value to
'   - valueToSet  - the value to assign to the item
'
' Usage:
'  AssignIfNull foo, bar
'
'-------------------------------------------------------------------------------
Sub AssignIfEmpty(byRef variableToSet, valueToSet) ' ASSIGNS IF THE CONTAINER IS EMPTY
   If IsNull(variableToSet) Then
      Assign variableToSet, valueToSet
   End If
   If IsEmpty(variableToSet) Then
      Assign variableToSet, valueToSet
   End If
   If NOT IsObject(variableToSet) Then
      If (variableToSet = "") Then
         Assign variableToSet, valueToSet
      End If
   End If
End Sub

'-------------------------------------------------------------------------------
' Sub: AssignIf
'
'  If the condition is true, do the assignment of the resultIfTrue, else resultIfFalse
'  see Notes for additional details
'
' Parameters:
'   - variableToAssign     - the variable to assign to
'   - condition            - the condition to evaluate
'   - resultIfTrue         - the value to assign if the condition evaluates to true
'   - resultIfFalse        - the value to assign if the condition evaluates to false
'
' Notes:
'  This shortcut replaces steps for checking both the result of the condition
'  AND whether the values to be assigned is an object or not
'-------------------------------------------------------------------------------
Sub AssignIf (ByRef variableToAssign, condition, resultIfTrue, resultIfFalse)
	On Error Resume Next
	Err.Clear
	Assign variableToAssign, Iif(condition, resultIfTrue, resultIfFalse)
	If Err.Number > 0 Then
		variableToAssign = False
	End If
	On Error Goto 0
End Sub


'============================================================================================================
' LOGIC

'-------------------------------------------------------------------------------
' Function: Iif
'
' A simplistic replacement for the ternary operater for in C or Perl
'
' Parameters:
'   - condition - The boolean condition to evaluate
'   - trueValue - Return this if the condition is true
'   - falseValue - Return this if the condition is false
'
' Returns:
' - variant - trueValue or falseValue depending on condition
'-------------------------------------------------------------------------------
Function Iif (condition, trueValue, falseValue)
   If condition Then
      Iif = trueValue
   Else
      Iif = falseValue
   End If
End Function

'============================================================================================================
' TYPING AND CONVERSION

'-------------------------------------------------------------------------------
' Function: IsAllBlank
'
'  Returns true if all the values in the array are NULL, Empty or "<blank>"
'
' Parameters:
'   - arrayOfValuesToCheck
'
' Returns:
'  true if all the values in the array are NULL, Empty or "<blank>"
'
'-------------------------------------------------------------------------------
Function IsAllBlank(arrayOfValuesToCheck)
   Dim result, loopResult
   result = False
   
   Dim i
   For i = 0 to UBound(arrayOfValuesToCheck)
      loopResult = True
      If (arrayOfValuesToCheck(i) = "<blank>") OR IsNull(arrayOfValuesToCheck(i)) OR IsEmpty(arrayOfValuesToCheck(i)) Then
         loopResult = False
      End If
      result = loopResult OR result
   Next
   
   IsAllBlank = result
ENd Function

'-------------------------------------------------------------------------------
' Sub: VerifyNotEmpty
'
'  if the variableToCheck is empty or NULL, logs a FATAL error
'
' Parameters:
'   - variableToCheck      - the variable to check the value of
'   - routineReporting     - routine that's reporting the error if this fails
'   - errorMessageToReport - the message to post in the event of it not being empty
'
'-------------------------------------------------------------------------------
Sub VerifyNotEmpty(variableToCheck,routineReporting, errorMessageToReport)
   If IsEmpty(variableToCheck) OR IsNull(variableToCheck) Then
      LogFatal "routine=>" & routineReporting & "|message=>" & errorMessageToReport
   End If
End Sub

'-------------------------------------------------------------------------------
' Sub: VerifyDataValue
'
'  if the variableToCheck is empty or NULL, logs a FATAL error
'
' Parameters:
'   - expressionToCheck    - the expression, should evaluate to either True or False
'   - routineReporting     - routine that's reporting the error if this fails
'   - errorMessageToReport - the message to post in the event of the expression evaluating to false
'
'-------------------------------------------------------------------------------
Sub VerifyDataValue (expressionToCheck, routineReporting, errorMessageToReport)
   If NOT expressionToCheck Then
      LogFatal "routine=>" & routineReporting & "|message=>" & errorMessageToReport
   End If
End Sub

'-------------------------------------------------------------------------------
' Sub: VerifyValueNotEmpty
'
'  if the variableToCheck is empty or NULL, logs a FATAL error
'
' Parameters:
'   - expressionToCheck    - the expression, should evaluate to either True or False
'   - routineReporting     - routine that's reporting the error if this fails
'   - errorMessageToReport - the message to post in the event of the expression evaluating to false
'
'-------------------------------------------------------------------------------
Sub VerifyValueNotEmpty(valueToCheck,errorNumber,errorMessageToReport)
	If isReallyEmpty(valueToCheck) Then
      LogFatal "routine=>VerifyValueNotEmpty|message=>" & errorNumber & " - " & errorMessageToReport
	End If
End Sub

'-------------------------------------------------------------------------------
' Function: MakeBool
'
'  Takes a string value and makes a boolean
'
' Parameters:
'  - incomingValue - value to be rendered as a boolean
'
' Returns:
'  True or False, depending on the input value (see notes)
'
' Notes:
'  Conditions that will generate a true result are anything EXCEPT:
'  False, Null, Empty, "", 0, "Off", "No", "False", "Blank", "0"
'-------------------------------------------------------------------------------
Function MakeBool(ByVal incomingValue)
	Dim result
	result = True
	AssignIf result, IsNull(incomingValue), False, result
	AssignIf result, incomingValue = False, False, result
	AssignIf result, incomingValue = Empty, False, result
	AssignIf result, Trim(incomingValue) = "", False, result
	If IsNumeric(incomingValue) Then
		AssignIf result, incomingValue = 0, False, result
	End If
	AssignIf result, Contains("OFF NO FALSE UNCHECKED 0 BLANK",UCase(incomingValue)), False, result

	MakeBool = result
End Function

'-------------------------------------------------------------------------------
' Function: CastAs
'
'  Let's me forget about all these functions to convert thing to thing
'  and you just give it the data and the target type as a string.
'
' Parameters:
'   - typeOfThingAsString - the string of the type you wanna make, such as "boolean"
'   - thingToCast         - the thing you wanna convert
'
' Returns:
'   - the converted thing
'
' Usage:
'   foo = CastAs("boolean",myVariable)
'
'-------------------------------------------------------------------------------
Function CastAs(typeOfThingAsString, thingToCast)
   Dim result
   Select Case LCase(typeOfThingAsString)
      Case "string"
         result = CStr(thingToCast)
      Case "s"
         result = CStr(thingToCast)
      Case "integer"
         result = CInt(thingToCast)
      Case "i"
         result = CInt(thingToCast)
      Case "int"
         result = CInt(thingToCast)
      Case "bool"
         result = MakeBool(thingToCast)
      Case "boolean"
         result = MakeBool(thingToCast)
      Case "b"
         result = MakeBool(thingToCast)
      Case Else
   End Select
   CastAs = result
End Function


'-------------------------------------------------------------------------------
' Function: IsDict
'
'  Returns true if item is an object and is a dictionary
'
' Parameters:
'   - incoming - item to test
'
' Returns:
'     True/False
'
' Usage:
'
'  If IsDict(widget) Then
'     ...
'
'-------------------------------------------------------------------------------
Function IsDict(incoming)
   Dim result
   result = False
   If IsObject(incoming) Then
      If NOT IsEmpty(incoming.Count) AND _
         NOT IsEmpty(incoming.Keys)  AND _
         NOT IsEmpty(incoming.Items) Then
            result = True
      End If
   End If
   IsDict = result
End Function

'-------------------------------------------------------------------------------
' Function: IsReallyEmpty
'
'  Returns true if the valueToCheck is NULL, Empty or ""
'
' Parameters:
'  - valueToCheck - the value to check
'
' Returns:
'  True if the thing is really empty
'
'-------------------------------------------------------------------------------
Function IsReallyEmpty(valueToCheck)
	Dim result
	result = FALSE
	On Error Resume Next
	If IsNull(valueToCheck) Then
		result = TRUE
	End If
	If IsEmpty(valueToCheck) Then
		result = TRUE
	End If
	If valueToCheck="" Then
		result = TRUE
	End If
	On Error Goto 0
	IsReallyEmpty = result
End Function

'-------------------------------------------------------------------------------
' Function: IsDefined
'
'  Returns true if the valueToCheck is NULL or Empty
'
' Parameters:
'  - valueToCheck - the value to check
'
' Returns:
'  True if the thing is really empty
'
'-------------------------------------------------------------------------------
Function IsDefined(valueToCheck)
   Dim result
   result = True
   If IsEmpty(valueToCheck) Then
      result = False
   End If
   If IsNull(valueToCheck) Then
      result = False
   End If
   IsDefined = result
End Function

'-------------------------------------------------------------------------------
' Function: MakeArrayIntoString
'
'  Takes the data in an array and attempts to render it as a string of human
'  readable characters. For instance and object gets replaced with "Object".
'  The intent is that this is used for debugging.
'
' Parameters: 
'  - theArray
'
' Returns:
'  A best guess string representation of the array
'
'-------------------------------------------------------------------------------
Function MakeArrayIntoString(theArray)
	MakeArrayIntoString = MakeItemPrintable(theArray)	' this is for backwards compatibility
End Function

'============================================================================================================
' "SMART DICT" SPECIFIC EXTENSIONS 

'-------------------------------------------------------------------------------
' Sub: DictMake
'
'  Renders incoming data into a dictionary. Takes arguments in the form of key/value 
'  pairs, seperated by => (key/value) or | (K/V pairs). Also takes arrays in the form 
'  Array(key1,value1,key2,value2...) 
'
' Parameters:
'   - Array or string of key/value pairs, as described above.
'   - Modifiers (also in dict format) or NULL
'
' Returns:
'   - Passed item is rendered as a dictionary
'
' Exceptions:
'   - None
'
' Usage:
'   MyRoutine "key1=>value1|key2=>value2" 
'   MyRoutine Array("key3",objectToPassIn,"key4",Array(1,2,3))
'   Sub MyRoutine args
'      DictMake args, NULL
'      print args("key1")
'
' Notes:
'  items prefaced with <eval> are passed to an eval function before being assigned, like so:
'  "boolean_thingie=><eval>True|array_thingie=><eval>Array(""A"",""B"")"
'  NOTE: THIS <eval> THING WILL ONLY WORK WITH LITERALS! To pass in an object, you have to 
'  use the array version of the calling convention
'-------------------------------------------------------------------------------
Sub DictMake (byref newDict, args)
   If NOT IsNull(args) AND NOT ISEmpty(args) Then
      DictMake args, NULL ' this has to be null to prevent the thing from recursing forever
   Else 
      Set args = CreateObject("Scripting.Dictionary")
   End If

   Dim passedStringArgs, passedArrayArgs, temp1, temp2
   ' were we passed an undefined item? if so, make it a dict
   If IsNull(newDict) OR IsEmpty(newDict) Then
      If args.Exists ("dict_make.new_object_call") Then
         Set newDict = Eval(args("dict_make.new_object_call"))
      Else
         Set newDict = CreateObject("Scripting.Dictionary")
         newDict.CompareMode = vbTextCompare
      End If
   Else

      If NOT IsObject(newDict) Then ' if it's an object by the time we get here, we're done
   
         If NOT IsString(newDict) AND NOT IsArray(newDict) Then
            newDict = Array("item",newDict)
         End If
   
         ' were we passed a string? if so, make it an array
         If NOT IsObject(newDict) AND NOT IsArray(newDict) Then
            ' then it must be string
            ' does it have a seperator in it?
            If InStr(1, newDict, "=>") = 0 Then
               newDict = "item=>" & newDict
            End If
            ' PAIR SEPS IS |, and KEY/VALUE SEPS ARE =>  
            passedStringArgs = newDict
            newDict = Replace(newDict,"=>","|") ' we're just going to convert it into a key/pair array
            newDict = Split(newDict,"|")
         End If
         
         ' and now we either have a dict, and nothing more to do
         ' OR we have an array, and need to make a dict. Sooo....
         If IsArray(newDict) Then
            ' The assumption is that even items in the array are keys, odd items are values
            passedArrayArgs = newDict
            Dim max
            max = UBound(newDict)
            If max => 0 Then  ' do we have any items?
               Dim i, result
               Set result = CreateObject("Scripting.Dictionary")
               For i = 0 to max step 2
                  If i+1 <= max Then ' do we have an odd number of items?
                     If IsObject(newDict(i+1)) Then   ' we have to deal with obejcts with Set
                        Set result(trim(newDict(i))) = newDict(i+1)
                     Else 
                        ' one more test, is this a string which should be eval'd?
                        If IsString(newDict(i+1)) Then
                           ' basic eval
                           If Left(CStr(newDict(i+1)),6)="<eval>" Then
                              newDict(i+1) = Eval(Mid(newDict(i+1),7))
                           End If
                           ' check for casting
                           If Left(CStr(newDict(i+1)),6)="<cast:" Then
                              temp1 = Mid(newDict(i+1),7)
                              temp1 = Mid(temp1,1,InStr(1,temp,">")-1)
                              temp2 = Mid(temp2,InStr(1,temp,">")+1)
                              newDict(i+1) = CastAs(temp1,temp2)
                           End If
                           
                        End If
                        result(Trim(newDict(i))) = newDict(i+1)
                     End If
                  Else ' yes, we have an odd number of items, render the value of the last one null
                     result(trim(newDict(i))) = NULL
                  End If
               Next
            End If
            Set newDict = result
            ' now we stash the raw data for debugging
            newDict("dictmake.rawarray") = passedArrayArgs
            If NOT IsEmpty(passedStringArgs) Then
               newDict("dictmake.rawstring") = passedStringArgs
            End If
         End If
         
      Else
         ' this code is to deal with the wacky case where you're debugging, and 
         ' some of the debugging code has made a temporary dict for you, which will
         ' go out of scope as soon as the subroutine ends.
         If newDict.Count = 0 Then
            If NOT (DictGet(args,"dict_make.skip_create_on_blank") = True) Then
               If args.Exists ("dict_make.new_object_call") Then
                  Set newDict = Eval(args("dict_make.new_object_call"))
               Else
                  Set newDict = CreateObject("Scripting.Dictionary")
                  newDict.CompareMode = vbTextCompare
               End If
            End If
         End If
      End If
   
   End If
   
End Sub

'-------------------------------------------------------------------------------
' Sub: DictMakeExpectItem
'
'  Renders incoming data into a dictionary. Takes arguments in the form of key/value 
'  pairs, seperated by => (key/value) or | (K/V pairs). Also takes arrays in the form 
'  Array(key1,value1,key2,value2...) 
'
'  NOTE: This function will return a dict containing the key:item. 
'
'  By calling this function, the caller is saying that what they want is a set
'  of args that contain a key:item. If the passed dict, after initial processing 
'  by DictMake DOES NOT contain a key:item, then the whole dict will be placed 
'  inside of another dict, and labelled as item. 
'
'  THIS SHOULD ONLY BE USED BY ROUTINES EXPECTING A key:item
'
' Parameters:
'   - Array or string of key/value pairs, as described above.
'   - Modifiers (also in dict format) or NULL
'
' Returns:
'   - Passed item is rendered as a dictionary
'
' Notes:
'  See DictMake
'-------------------------------------------------------------------------------
Function DictMakeExpectItem (byref newDict, args)
   Dim result, temp
   DictMake newDict, args
   set result = newDict
   If NOT newDict.Exist("item") Then
      Set temp = DictCreate(NULL)
      Set temp("item") = result
      temp("this_object_auto_generated") = True
      Set result = temp
   End If
   Set newDict = result
   Set DictMakeExpectItem = result
End Function

'-------------------------------------------------------------------------------
' Function: DictUnwrapItemIfExists
'
'  Undoes DictMakeExpectItem
'
'  THIS SHOULD ONLY BE USED BY ROUTINES EXPECTING A key:item
'
' Parameters:
'   - dict which was processed by DictMakeExpectItem
'
' Returns:
'   - args("item")
'
' Notes:
'  See DictMakeExpectItem
'-------------------------------------------------------------------------------
Function DictUnwrapItemIfExists (args)
   If args.Exist("this_object_auto_generated") Then
      Set args = args("item")
   End If
   Set DictUnwrapItemIfExists = args
End Function

'-------------------------------------------------------------------------------
' Function: DictCreate
'
'  Returns a dict from some passed args
'
' Parameters:
'   - args              - a DictMake compatible input
'
' Returns:
'   - a dictionary with the keys/values added
'
' Usage:
'   Dim foo
'   Set foo = DictCreate("key1=>value1|key2=>value2")
'
'-------------------------------------------------------------------------------
Function DictCreate(args)
   DictMake args, NULL
   Set DictCreate = args
End Function

'-------------------------------------------------------------------------------
' Sub: DictCopy
'
'  Copies keys from one dict to another. performs shallow copy.
'
' Parameters (required):
'   - args              - a DictMake compatible input
'   - key:from          - dictionary: to copy from
'   - key:to            - dictionary: to copy to
'
' Parameters (allowed):
'   - key:keys_to_copy  - array: if null, just copy all the keys from the source dict
'   - key:keys_to_avoid - array: avoid copying these keys
'
' Returns:
'   - renders both dicts as dicts via DictMake
'
' Usage:
'   DictCopy Array("to",dictToCopyTo,"from",dictToCopyFrom)
'   DictCopy Array("to",dictToCopyTo,"from",dictToCopyFrom, "keys_to_copy",Array("key1","key2","key3"))
'
'-------------------------------------------------------------------------------
Sub DictCopy (copyFrom, copyTo, args)

   DictMake args, NULL
   DictMake copyFrom, NULL
   DictMake copyTo, NULL
   
   Dim keysToCopy
   Dim keysToAvoid
   Dim passedKeysToAvoid

   ' were we given a list of keys to copy?
   keysToCopy = DictWithdraw (args, "keys_to_copy")
   
   If IsEmpty(keysToCopy) Then
      KeysToCopy = copyFrom.Keys
   End If
   If IsString(KeysToCopy) Then
      KeysToCopy = Array(KeysToCopy)
   End If

   Set keysToAvoid = CreateObject("Scripting.Dictionary")
   keysToAvoid.CompareMode = vbTextCompare
   keysToAvoid("dictmake.rawarray") = True
   keysToAvoid("dictmake.rawstring") = True
   passedKeysToavoid = DictWithdraw(args, "keys_to_avoid")
   If IsString(passedKeysToAvoid) Then
      passedKeysToAvoid = Array(passedKeysToAvoid)
   End If
   
   If IsArray(passedKeysToAvoid) Then
      If UBound(passedKeysToAvoid) => 0 Then
         Dim j
         For j = 0 to UBound(passedKeysToAvoid)
            keysToAvoid(passedKeysToAvoid(j)) = True
         Next
      End If
   End If
   
   If UBound(KeysToCopy) => 0 Then
      Dim i, selectedKey
      For i = 0 to UBound(KeysToCopy)
         ' TBD: ADD CODE HERE TO REJECT ITEMS FROM THE KEYS TO AVOID
         Assign selectedKey, KeysToCopy(i)
         If NOT keysToAvoid(selectedKey) = True Then
            If IsObject(copyFrom(selectedKey)) Then
               Set copyTo(selectedKey) = copyFrom(selectedKey)
            Else
               copyTo(selectedKey) = copyFrom(selectedKey)
            End If
         End If
      Next
   End If
   
End Sub

'-------------------------------------------------------------------------------
' Function: DictGet
'
'  Fetches a value from a dictionary if the value already exists. Will not 
'  create an empty key on access the way dict("foo") does.
'
' Parameters:
'   - dictionaryToUse - the dictionary to attempt to fetch a value from
'   - keyToCheckFor   - the key to check existance of/fetch value from
'
' Returns:
'  Either Empty or the value found in the dictionary
'
' Usage:
'  myValue = DictGet(theDict, "foo")
'
' Notes:
'  We need this because the simply using if dict("key") results in an empty key being created.
'
'-------------------------------------------------------------------------------
Function DictGet(dictionaryToUse, keyToCheckFor)
   DictGet = Empty
   If dictionaryToUse.Exists (keyToCheckFor) Then
      If IsObject(dictionaryToUse(keyToCheckFor)) Then
         Set DictGet = dictionaryToUse(keyToCheckFor)
      Else
         DictGet = dictionaryToUse(keyToCheckFor)
      End If
   End If
End Function

'-------------------------------------------------------------------------------
' Function: DictWithdraw
'
'  Returns the specified value, and deletes the value from the dictionary.
'  Non existant keys return empty and do not create empty keys.
'
' Parameters:
'   - dictionaryTouse - Must *ALREADY* be a dictionary
'   - keyToCheckFor   - Key to select
'
' Returns:
'  The item found at that key or Empty if no key found
'
' Usage:
'  DictMake args, NULL
'  mySSN = DictWithdraw args, "SSN"
'  theOtherDict.ApplyKeys (args) ' copies everything *still* in args to theOtherDict
'
'-------------------------------------------------------------------------------
Function DictWithdraw(dictionaryTouse, keyToCheckFor)
   DictWithdraw = Empty
   If dictionaryToUse.Exists (keyToCheckFor) Then
      If IsObject(dictionaryToUse(keyToCheckFor)) Then
         Set DictWithdraw = dictionaryToUse(keyToCheckFor)
      Else
         DictWithdraw = dictionaryToUse(keyToCheckFor)
      End If
      dictionaryTouse.Remove keyToCheckFor
   End If
End Function


'-------------------------------------------------------------------------------
' Function: MakeDictIntoArray
'
'  Render a Dictionary into an Array. Form is key,value,key,value...
'
' Parameters:
'   - inDictionary - the dictionary to render as an array
'
' Returns:
'  An arry with the contents of the dict in the form key,value,key,value...
'-------------------------------------------------------------------------------
Function MakeDictIntoArray(inDictionary)

   dim result(), loopi, keyArray, itemArray
   redim result( (inDictionary.count *2) - 1 )
   dim aKey, anItem
   if  inDictionary.count = 0 then
      MakeDictIntoArray = empty
      exit Function  '<<<
   end if
   keyArray = inDictionary.keys
   itemArray = inDictionary.items

   for loopi = 0 to inDictionary.count-1
      result(loopi*2) = keyArray(loopi) 
      assign result(loopi*2 + 1),  itemArray(loopi)
   next

   MakeDictIntoArray = result
end function

'-------------------------------------------------------------------------------
' Sub: GlobalDictionaryAdd
'
'	This function is here because I don't want to have to always be checking existance 
'	in order to set defaults... Like TestName. Reduces 4 lines of code to 1
'
' Parameters:
'   - KeyName - String - Name of the key to add or overwrite
'   - KeyValue - anytype - Data to add
'
'-------------------------------------------------------------------------------
Sub GlobalDictionaryAdd (KeyName,KeyValue)
	If GlobalDictionary.Exists(KeyName) Then
		GlobalDictionary.Remove(KeyName)
	End If
	GlobalDictionary.Add KeyName, KeyValue
End Sub

'-------------------------------------------------------------------------------
' Sub: GlobalDictionaryRemove
'
'	If a value is defined, removes it
'
' Parameters:
'   - KeyName - String - Name of the key to add or overwrite
'
'-------------------------------------------------------------------------------
Sub GlobalDictionaryRemove (KeyName)
	If GlobalDictionary.Exists(KeyName) Then
		GlobalDictionary.Remove(KeyName)
	End If
End Sub

'-------------------------------------------------------------------------------
' Sub: DictionaryAdd
'  Adds a value to a dictionary. removes that key first if it's already there
'
' Parameters:
'   - dictionaryToUse - the dictionary to add to (can be uninitialized)
'   - keyName         - the key to add the value under
'   - keyValue        - the value to add
'-------------------------------------------------------------------------------
Sub DictionaryAdd (dictionaryToUse,KeyName,KeyValue)
   DictMake dictionaryToUse, NULL
	DictionaryRemove dictionaryToUse, KeyName
	dictionaryToUse.Add KeyName, KeyValue
End Sub

'-------------------------------------------------------------------------------
' Sub: DictionaryAddConditional 
'  If the specified key does not already exist, add it. 
'  If it DOES exist, do nothing
'  Creates the dict if the passed dictionaryToUse hasn't been initialized
'
'  Used in the creation of complex records-containing-other-records
'  where we use the dict as the record
'
' Parameters:
'   - dictionaryToUse - the dictionary to use (can be uninitialized)
'   - newKey          - key to conditionally add
'   - newValue        - the new value to add
'
' Usage:
'  DictionaryAddConditional memberRecord("coverageTypeSpanRecord"), "EnrollmentType", "New Hire"
'-------------------------------------------------------------------------------
Sub DictionaryAddConditional(dictionaryToUse,newKey,newValue)
   DictMake dictionaryToUse, NULL
   If IsNull(dictionaryToUse) OR IsEmpty(dictionaryToUse) Then
      Set dictionaryToUse = CreateObject("Scripting.Dictionary")
   End If
   If NOT dictionaryToUse.Exists(newKey) Then
      dictionaryToUse.Add newKey, newValue
   End If
End Sub


'-------------------------------------------------------------------------------
' Sub: DictionaryRemove
'  Removes a value from a dictionary. (trying to remove a key that isn't there
'  causes an exception, so this does the check first)
'
' Parameters:
'   - dictionaryToUse - the dictionary to use
'   - keyName         - key to remove
'-------------------------------------------------------------------------------
Sub DictionaryRemove (dictionaryToUse, KeyName)
   DictMake dictionaryToUse, NULL
	If dictionaryToUse.Exists(KeyName) Then
		dictionaryToUse.Remove(KeyName)
	End If
End Sub

'-------------------------------------------------------------------------------
' Function: DictionaryMake
'  DictMake is a sub, this is a complimentary function. returns the item 
'  from the function. same functionality, differnet calling form than DictMake
'
' Parameters:
'   - itemsToPopulateWith  - a DictMake compatible list of arguments
'
' Returns:
'	- Created object
'
' Exceptions:
'   - None
'-------------------------------------------------------------------------------
Function DictionaryMake (itemsToPopulateWith)
   Dim result
   Assign result, itemsToPopulateWith
   DictMake result, NULL
   if IsNull(result) OR IsEmpty(result) Then
      Set result = CreateObject("Scripting.Dictionary")
   End If
   Set DictionaryMake = result
End Function

'-------------------------------------------------------------------------------
' Function: ReadPrefsFileIntoDict
'
'  reads a unix style preferences file and puts the name/value pairs into a dictionary
'
' Parameters:
'   - fileName - name of the file to read from
'
' Returns:
'  dictionary of name/value pairs populated from the file
'
' Notes:
'  DEPENDENCY ON class-lists.vbs
'-------------------------------------------------------------------------------
Function ReadPrefsFileIntoDict(fileName)
   Dim result
   Set result = CreateObject("Scripting.Dictionary")
   Dim interimData
   Set interimData = NewList()
   interimData.ReadFromFile(fileName)
   Dim indexIntoPairs, temp
   For indexIntoPairs = 0 to interimData.MaXindex
      temp = InStr(1, interimData.Item(indexIntoPairs), "=")
      DictionaryAdd result, Mid(interimData.Item(indexIntoPairs),1,temp-1), Mid(interimData.Item(indexIntoPairs),temp+1)
   Next
   Set ReadPrefsFileIntoDict = result
End Function

'-------------------------------------------------------------------------------
' Sub: CopyKeysBetweenDicts
'
'  used to copy a key or group of keys from one dict to another
'
' Parameters:
'   - dictionaryToCopyFrom - the source dictionary
'   - dictionaryToCopyTo   - the target dictionary
'   - arrayOfKeys          - the keys to copy
'-------------------------------------------------------------------------------
Sub CopyKeysBetweenDicts(dictionaryToCopyFrom, dictionaryToCopyTo, arrayOfKeys)
   Set dictionaryToCopyTo = DictionaryMake(dictionaryToCopyTo)
   dim i
   For i = 0 to UBound(arrayOfKeys)
      dictionaryToCopyTo(arrayOfKeys(i)) = dictionaryToCopyFrom(arrayOfKeys(i))
   Next
End Sub

'-------------------------------------------------------------------------------
' Sub: RecordModify
'
'  This is used to modify dictionaries that are used as records, and may contain 
'  other similar dictionaries. 
'
' Parameters:
'   - recordToModify         - the dictionary object
'   - arrayOfRecordParantage - the array of keys to the specific subrecord
'   - fieldName              - the field name for the key to modify
'   - newValue               - the new value for that key
'
' Notes:
'  This was developed for testing a record management application 
'  where each record could have many sub records and sub-sub records
'-------------------------------------------------------------------------------
Sub RecordModify(byRef recordToModify, arrayOfRecordParantage,fieldName,newValue)

   Dim dq, i, leftitem, rightitem, theRecord
   
   dq = chr(34)
   If IsNull(recordToModify) OR IsEmpty(recordToModify) Then
      Set recordToModify = CreateObject("Scripting.Dictionary")
   End If

   i = 0
   leftitem = "recordToModify"
   ' arrayOfRecordParantage is null? means just add the field to the record
   If IsNull(arrayOfRecordParantage) Then
      If recordToModify.Exists(fieldName) Then
         recordToModify.Remove fieldName
      End If
      recordToModify.Add fieldName, newValue
   Else

      ' is string? make into array
      If NOT IsArray(arrayOfRecordParantage) Then
         arrayOfRecordParantage = Array(arrayOfRecordParantage)
      End If

      ' build the tree
      If UBound(arrayOfRecordParantage) > 0 Then
         ' arrays of more than one element
         For i = 0 to UBound(arrayOfRecordParantage)
      
            rightitem = dq & arrayOfRecordParantage(i) & dq
            'print leftitem & ".Exists(" & rightitem & ") = " & Eval(leftitem & ".Exists(" & rightitem & ")")
            If NOT Eval(leftitem & ".Exists(" & rightitem & ")") Then
               Eval(leftitem).Add arrayOfRecordParantage(i), CreateObject("Scripting.Dictionary")
            End If
            'print leftitem & ".Exists(" & rightitem & ") = " & Eval(leftitem & ".Exists(" & rightitem & ")")
            leftitem = leftitem & "(" & rightitem & ")"
            'print "new leftitem = " & leftitem
         Next
      Else
         'arrays of one element
         If NOT recordToModify.Exists(arrayOfRecordParantage(0)) Then
            recordToModify.Add arrayOfRecordParantage(0), CreateObject("Scripting.Dictionary")
         End If
         leftitem = "recordToModify" & InPeren(InQuotes(arrayOfRecordParantage(0)))
      End If
   
      Set theRecord = Eval(leftitem)
      If theRecord.Exists(fieldName) Then
         theRecord.Remove fieldName
      End If   
      theRecord.Add fieldName, newValue
      'print leftitem & "(" & dq & fieldName & dq & ")=" & theRecord(fieldName)

   End If

End Sub

'-------------------------------------------------------------------------------
' Function: MakeCopyOfDict
'
'  This does a copy of one dict to another. It does a simple array from the
'  keys and values, so in VBScript this is a shallow copy
'
' Parameters:
'   - dictionaryToCopy         - the dictionary to copy
'-------------------------------------------------------------------------------
Function MakeCopyOfDict(dictionaryToCopy)
   Dim result
   result = MakeDictIntoArray(dictionaryToCopy)
   DictMake result, NULL
   Set MakeCopyOfDict = result
End Function

'-------------------------------------------------------------------------------
' Function: IsAllKeysBlank
'
'  Evaluates all the specified entries in a dict to see if they're all blank
'
' Parameters:
'   - dictionaryToCheck         - the dictionary to copy
'   - arrayOfKeys               - array of keys to check to see if they're blank
'
'-------------------------------------------------------------------------------
Function IsAllKeysBlank(dictionaryToCheck,arrayOfKeys)
   Dim result,i
   result = True
   For i = 0 to UBound(arrayOfKeys)
      If NOT IsReallyEmpty(dictionaryToCheck(arrayOfKeys(i))) Then
         result = False
      End If
   Next
   IsAllKeysBlank = result
End Function

'===============================================================================
' OTHER OBJECT MANIPULATION

'-------------------------------------------------------------------------------
' Function: GetClass
'
' Parameters:
'   - incomingObject - the object who's class you wish to fetch
'
' Returns:
'	The class name for classes which contain a className property
'	else an empty string
'
'-------------------------------------------------------------------------------
Function GetClass(incomingObjectToGetClassOf) ' failed=incomingObject
	Dim classNameResult, className
	classNameResult = ""
	className = Empty

	On Error Resume Next
   Err.Clear

	classNameResult=incomingObjectToGetClassOf.className ' if we defined it, it should have this
   'If Err.Number > 0 Then
   '   LogWarning "routine= Get_cursor;message=Expected error encountered while attempting to fetch class: " & Err.Number & ", " & Err.Description
   '   Err.Clear
   'End If

	If classNameResult = "" Then

		' i guess we didn't... is it a window?
		If incomingObjectToGetClassOf.Exist(0) Then
			className = incomingObjectToGetClassOf.GetROProperty("nativeclass")
'         If Err.Number > 0 Then
'            LogWarning "routine=GetClass;message=Expected error encountered while attempting to fetch nativeclass: " & Err.Number & ", " & Err.Description
'            Err.Clear
'         End If

			If NOT IsEmpty(className) Then
				classNameResult = className
			End If
		Else
			classNameResult = "unable to determine object type"
		End If

	End If
	Err.Clear
	On Error Goto 0

	GetClass=classNameResult
End Function

'-------------------------------------------------------------------------------
' Function: GetItemIndexFromObjectContent
'
' Searches through the content of an object (as returned by the QTP GetContent
' mthod) for a given search item
'
' Parameters:
'   - searchItem: the string we're looking for
'   - objectContent: the content of the object as returned by the GetContent method,
'                    i.e. a single string delimited by vbLf characters
'
' Returns:
'  - integer - The index of the item in the content if found, else -1
'-------------------------------------------------------------------------------
Public Function GetItemIndexFromObjectContent (searchItem, objectContent)
   Dim items
   Dim i
   
   GetItemIndexFromObjectContent = -1 
   items = Split(objectContent, vbLf)
   
   For i = 0 to UBound(items)
      If items(i) = searchItem Then
         GetItemIndexFromObjectContent = i
         Exit For
      End If
   Next
End Function



'***********************************************************
' aggregating q:\utils\libs\extensionsstrings.vbs

'###############################################################################
' Library: ExtensionsStrings
'
' About: Basic extensions to VBScript/QTP - Akien MacIain
'  Things that *should* have been in VBScript gathered together
'  Functions supporting working with strings
'
'  Copyright (C) 2008, 2009, 2010 Akien MacIain
'
'  This program is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'###############################################################################
'Option Explicit

'===============================================================================
' Section: Public Functions
'   Functions that are exported by the library
'===============================================================================

'-------------------------------------------------------------------------------
' Function: FrameworkDetectExtensionsStrings
'   Utility function for the Framework Compilation checking utility
'
' Returns:
'  (integer) always returns 1
'-------------------------------------------------------------------------------
Function FrameworkDetectExtensionsStrings()
	FrameworkDetectExtensionsStrings = 1
End Function


'-------------------------------------------------------------------------------
' Sub: QuickSort 
'   Implements a QuickSort for numeric or string arrays
'
' Parameters:
'  - sortArray: (array) A one dimensional array to be sorted
'  - loBound: (integer) The lower boundry of the array
'  - hiBound: (integer) The upper boundry of the array
'
'  Notes:
'   the arrya is sorted in place.  An example usage would be
'
'     (code)
'     QuickSort myArray, LBound(myArray), UBound(myArray)
'     (end code)
'
'  This code is taken from http://4guysfromrolla.com/webtech/012799-2.shtml.
'
'-------------------------------------------------------------------------------
Sub QuickSort(sortArray,loBound,hiBound)
   Dim pivot
   Dim loSwap
   Dim hiSwap
   Dim temp

   'Two items to sort
   if hiBound - loBound = 1 then
      if sortArray(loBound) > sortArray(hiBound) then
         temp = sortArray(loBound)
         sortArray(loBound) = sortArray(hiBound)
         sortArray(hiBound) = temp
      End If
   End If

   'Three or more items to sort
   pivot = sortArray(int((loBound + hiBound) / 2))
   sortArray(int((loBound + hiBound) / 2)) = sortArray(loBound)
   sortArray(loBound) = pivot
   loSwap = loBound + 1
   hiSwap = hiBound

   do
      'Find the right loSwap
      while loSwap < hiSwap and sortArray(loSwap) <= pivot
         loSwap = loSwap + 1
      wend
      'Find the right hiSwap
      while sortArray(hiSwap) > pivot
         hiSwap = hiSwap - 1
      wend
      'Swap values if loSwap is less then hiSwap
      if loSwap < hiSwap then
         temp = sortArray(loSwap)
         sortArray(loSwap) = sortArray(hiSwap)
         sortArray(hiSwap) = temp
      End If
   loop while loSwap < hiSwap

   sortArray(loBound) = sortArray(hiSwap)
   sortArray(hiSwap) = pivot

   'Recursively call function .. the beauty of Quicksort
   '2 or more items in first section
   if loBound < (hiSwap - 1) then QuickSort sortArray,loBound,hiSwap-1
   '2 or more items in second section
   if hiSwap + 1 < hibound then QuickSort sortArray,hiSwap+1,hiBound
End Sub

'-------------------------------------------------------------------------------
' Function: F
'  2008 Akien MacIain
'  Formats a string, a la printf  [*DEPRECATED* - Replaced by <Fmt>]
'
' Parameters:
'   - theData - string - formatting template
'   - passedArgs - array - items to replace with
'
' Returns:
'   - string with substitutions
'
' Notes:
'  is also aware of \n
'  F("%1 is %2",Array("one", 1)
'     %1
'  The following are planned, but not yet supported
'     %{d=1;w=17;j=l;p=0}
'     %{data=1;width=17;justify=l;padchar=0}
'-------------------------------------------------------------------------------
Function F(byval theData, byref passedArgs)
   If NOT IsArray(passedArgs) Then
      passedArgs=Array(passedArgs)
   End If
   Dim i
   If InStr(1,theData,"%") Then
      For i = 0 to UBound(passedArgs) 
         theData = Replace(theData,"%"&i,passedArgs(i))
      Next
   End If
   If InStr(1,theData,"\n") Then
      theData = Replace(theData,"\n",VbCrLf)
   End If
   If InStr(1,theData,"\d") Then
      theData = Replace(theData,"\d",Chr(34))
   End If
   If InStr(1,theData,"%{") Then
      raise 1, "That's not supported yet!"
   End If
   f = theData
End Function

'-------------------------------------------------------------------------------
' Function: Fmt
'  2008 Akien MacIain
'  Formats a string, a la printf
'
' Parameters:
'   - theData - string - formatting template
'   - passedArgs - array - items to replace with
'
' Returns:
'   - string with sustitutions
'
' Notes:
'  is also aware of \n
'  F("%1 is %2",Array("one", 1)
'     %1
'  The following are planned, but not yet supported
'     %{d=1;w=17;j=l;p=0}
'     %{data=1;width=17;justify=l;padchar=0}
'-------------------------------------------------------------------------------
Function Fmt(byval formatString, byref passedArgs)
   If NOT IsArray(passedArgs) Then
      passedArgs=Array(passedArgs)
   End If
   Dim i
   If InStr(1,formatString,"%") Then
      For i = 0 to UBound(passedArgs) 
         formatString = Replace(formatString,"%"&i,passedArgs(i))
      Next
   End If
   If InStr(1,formatString,"\n") Then
      formatString = Replace(formatString,"\n",VbCrLf)
   End If
   If InStr(1,formatString,"\d") Then
      formatString = Replace(formatString,"\d",Chr(34))
   End If
   If InStr(1,formatString,"%{") Then
      raise 1, "That's not supported yet!"
   End If
   Fmt = formatString
End Function

'-------------------------------------------------------------------------------
' Function: CArrayOfStringFromDict
'
'  takes a dict or a bag_dict and renders them as an array of string 
'  using the DictMake way of displaying the information. 
'
'  Primary use is object inspection at run time during debugging
'
' Parameters:
'   - dictToReturn - can be a dict or a bag_dict (or anything implementing those interfaces)
'
' Returns:
'  array of string in the form Array("key1=>value1","key2=>value2",...)
'
' Exceptions:
'   - none. in case of error, returns string with error message (again, for clarity in debugging)
'
' Usage:
'  In the debug watch window: CArrayOfStringFromDict(myDict)
'
'-------------------------------------------------------------------------------
Function CArrayOfStringFromDict(dictToReturn)

   Dim result, i, theKeys, theValues

   If NOT IsObject(dictToReturn) Then
      result = "passed an item which is not a dict"
   Else

      Set result = CreateObject("Scripting.Dictionary")
      theKeys = dictToReturn.Keys
      theValues = dictToReturn.Items
      
      For i = 0 to UBound(theKeys)
         result( theKeys(i) & "=>" & CString(theValues(i)) ) = True 
      Next
   
      result = result.Keys

   End If

   CArrayOfStringFromDict = result
End Function 

'-------------------------------------------------------------------------------
' Function: CString
'
'  Attempts to take whatever is passed to it and render it as a string
'  NULL gets rendered as <null>, Empty gets rendered as <Empty>
'  Arrays are rendered as string representations of arrays.
'
'  For objects of any kind, the code uses a series of rules rendered as
'  If statements to attempt to determine useful information about the object
'  and return class, window, dictionary or bag object information.
'  For most of the classes we've defined, we've implemented a
'  ClassName property. This will attempt to detect that. 
'
' Parameters:
'   - itemToRenderAsString
'
' Returns:
'  either "" or a string representation of whatever was passed
'
' Exceptions:
'   - none in this code, tho in the final else, anything that simply cannot
'     be rendered as a string could well cause errors. THOSE ARE 
'     DELIBERATELY NOT TRAPPED.
'
' Usage:
'  myString = CString(someOtherThingie)
'
'-------------------------------------------------------------------------------
Function CString(byVal itemToRenderAsString)
   Dim result, i
   result = ""
   If IsNull(itemToRenderAsString) Then
      result = "<null>"
   ElseIf IsEmpty(itemToRenderAsString) Then
      result = "<empty>"
   ElseIf IsArray(itemToRenderAsString) Then
      For i = 0 to UBound(itemToRenderAsString)
         result = result & ", " & CString(itemToRenderAsString(i))
      Next
      result = "Array(" & Mid(result,2) & ")"
   ElseIf IsObject(itemToRenderAsString) Then
      result = "Unknown object"
      If IsArray(itemToRenderAsString.keys) AND IsArray(itemToRenderAsString.items) Then
         result = "dictionary object"
      End If
      If itemToRenderAsString.Exists("self.is_bag_object") Then
         result = "bag_object of class " & itemToRenderAsString("self.class_name")
      End If
   Else
      result = "" & itemToRenderAsString & ""
   End If
   CString = result
End Function

'-------------------------------------------------------------------------------
' Function: MakeDictIntoString
'  Render a dictionary as a string for debugging [*DEPRECATED* - replaced by <CString>]
'
' Parameters:
'   - inDictionary - incoming dictionary
'
' Returns:
'  Contents of dictionary rendered as a set of strings
'
' Notes:
'   This is an older form of CString (See ExtensionsStrings.vbs)
'-------------------------------------------------------------------------------
Function MakeDictIntoString(inDictionary)

   dim result, loopi, keyArray
   dim aKey, anItem
   keyArray = inDictionary.keys
   for each akey in keyArray
      result = result & akey & "="
      assign anItem, inDictionary.item(akey)
      if isArray(anItem) then
			for loopi = 0 to ubound(anItem)
            result = result & anItem(loopi) & ","
			next
         result = result & ";"
      elseif isObject(anItem) then
         result = result & "is object;"
      else
         result = result & anItem & ";"
      end if
   next

   MakeDictIntoString = result
end function
'-------------------------------------------------------------------------------
' Function: IsString
'
'  Returns true if item is a string
'
' Parameters:
'   - incoming - item to test
'
' Returns:
'     True/False
'
' Usage:
'  
'     If IsString(foo) Then
'        ...
'
'-------------------------------------------------------------------------------
Function IsString(incoming)
   IsString = (TypeName(incoming) = "String")
End Function

'-------------------------------------------------------------------------------
' Function: Between
' returns the string found between two other strings
'
' Parameters:
'  - superstring: (string) the string to search 
'  - leftstring: (string) the left string bounding the string we're looking for
'  - rightstring: (string) the right string bounding the string we're looking for 
'-------------------------------------------------------------------------------
Function Between(superstring, leftstring, rightstring)
	Dim result, iLeftIndex, iRightIndex
	result = ""
	iLeftIndex=Instr(1, superstring, leftstring,1)+Len(leftstring)
	iRightIndex=Instr(iLeftIndex,superstring,rightstring,1)
	If iRightIndex>iLeftIndex Then
		result=Mid(superstring,iLeftIndex,iRightIndex-iLeftIndex)
	End If
	Between=result
End Function

'-------------------------------------------------------------------------------
' Function: Contains
' Returns True/False of does the first arg contain the second arg?
'
' Parameters:
'   - superItem: The item to check in
'   - subItem: the item to check for
'
' Returns:
'  - boolean
'-------------------------------------------------------------------------------
Function Contains(superItem, subItem)
	Dim result
	result = False

	' currently the only support is for strings. 
	' here's where we'd add type checking for other data types
	' i envision we might need to be able to process a List object
	' or date ranges, or validate against lists within the UI of the AUT
	If InStr(1,superItem,subItem) > 0 Then
		result = True
	End If
	
	Contains = result
End Function

'-------------------------------------------------------------------------------
' Function: IsNumbersOnly
'   Checks if a string contains only numbers (0-9)
'
' Parameters:
'   - Input: (String) The string to check
'
' Returns:
'  (Boolean) True if the string contains only numbers, false if it contains
'  anything else.
'
'-------------------------------------------------------------------------------
Function IsNumbersOnly(Input)
	Dim numbers
	Dim i 
	numbers="0123456789"

   IsNumbersOnly = True
   
	For i = 1 to Len(Input)
		If not (InStr(1,numbers,Mid(Input,i,1))>0) Then
			IsNumbersOnly = False
         Exit For
		End If
	Next
End Function

'-------------------------------------------------------------------------------
' Function: Digits
'
'  Returns the numeric part of a string... $123,456.78 retuens 123456.78
'  because of the special nature of the - and . characters, "extras" are ignored
'  so 123-456-7890 becomes -1234567890, and ..1..2..3..4 becomes 1.234
'
' Parameters:
'   - instring - the string to transform
'
' Returns:
'  the modified string of digits
'
'-------------------------------------------------------------------------------
Function Digits(instring)
   Dim result, fMinus, fDecimal, sMasterList, i, s, iPosition
	result=""
	fMinus=FALSE
	fDecimal=FALSE
	sMasterList="--0123456789."
	' walk the string, character by character...
	For i = 1 to len(instring)
		s = Mid(instring,i,1)
		iPosition=InStr(1,sMasterList,s)
		If iPosition>0 Then	' does it appear in the valid characters list?
			If (s="-") Then
			' we can only have 1 negetion symbol, and it must go at the beginning
				If fMinus=FALSE Then
					fMinus=TRUE
					result="-"&result
				End If
			ElseIf s="." Then
				' we can only have one decimal symbol
				If fDecimal=FALSE Then
					fDecimal=TRUE
					result=result&s
				End If
			Else
				result=result&s		' not a - or a . but does appear in the valid list, so add the character!
			End If
		End If
	Next
	If result="" Then
		result=0
	End If
	Digits=result
End Function

'-------------------------------------------------------------------------------
' Function: StrictDigits
'
'  returns just the digits (pays no attention to "-" or ".")
'
' Parameters:
'  - incomingString - the string to transform
'
' Returns:
'  the transofmred string
'
'-------------------------------------------------------------------------------
Function StrictDigits(incomingString)
	Dim allowedCharacters
	Dim resultString
	Dim i 
	allowedCharacters="0123456789"
	resultString=""
	
	For i = 1 to Len(incomingString)
		If InStr(1,allowedCharacters,Mid(incomingString,i,1))>0 Then
			resultString = resultString & Mid(incomingString,i,1)
		End If
	Next

	StrictDigits=resultString
End Function

'-------------------------------------------------------------------------------
' Function: twoDigits
'
'  returns Right("0" + incomingString,2)
'
' Parameters:
'  - incomingString - the string to transform
'
' Returns:
'  the transofmred string
'
'-------------------------------------------------------------------------------
Function twoDigits(incomingString)
   incomingString = "0" & Digits(incomingString)
   incomingString= Right(incomingString,2)
   twoDigits=incomingString
End Function

'-------------------------------------------------------------------------------
' Sub: SimpleSortStringArray
'
'  Takes the passed array and returns it sorted
'
' Parameters:
'   - sortMe - arry of strings to sort
'
' Returns:
'   Modifies the passed array
'-------------------------------------------------------------------------------
Sub SimpleSortStringArray(ByRef sortMe)

	Dim keepGoing, changeHappened, tempValue, loopCounter
	keepGoing = True
	
	While keepGoing

		changeHappened = False

		For loopCounter = 0 to UBound(sortMe)-1
			If (LCase(sortMe(loopCounter)) > LCase(sortMe(loopCounter+1))) OR IsEmpty(sortMe(loopCounter)) OR (sortMe(loopCounter)="") OR IsNull(sortMe(loopCounter))Then
				tempValue = sortMe(loopCounter)
				sortMe(loopCounter) = sortMe(loopCounter+1)
				sortMe(loopCounter+1) = tempValue
				changeHappened = True
			End If
		Next

		If NOT changeHappened Then
			keepGoing = False
		End If

	Wend


End Sub

'-------------------------------------------------------------------------------
' Function: InQuotes
'
'  shortcut to wrap a string in quotes. (Written before I knew how to escape them)
'
' Parameters:
'   - stringToUse - the string to wrap in quotes
'
' Returns:
'  The string wrapped in quotes
'
'-------------------------------------------------------------------------------
Function InQuotes(stringToUse)
   InQuotes = chr(34) & stringToUse & chr(34)
End Function

'-------------------------------------------------------------------------------
' Function: InPeren
'
'  shortcut to wrap a string in parenthesis
'
' Parameters:
'   - stringToUse - the string to wrap in parenthesis
'
' Returns:
'  The string wrapped in parenthesis
'
'-------------------------------------------------------------------------------
Function InPeren(stringToUse)
   InPeren = "(" & stringToUse & ")"
End Function

'-------------------------------------------------------------------------------
' Function: RoundString
'
'  Used to take a string of text, extract the digits, and round it
'  So can take "You will make #123,456.78.9 dollars" and turn it into 123456.79
'
' Parameters:
'   - sIncoming   - string to transform
'   - iDecPlaces  - number of decimal places to round to
'
' Returns:
'  The resultant number, after non digits are removed and it's rounded
'
' Notes:
'  The last time I (Akien) worked on this function, I realized it was no longer
'  being used by any part of the framework. I am leaving it in because it was used
'  at one point, it's already here, and is potentially useful in the future
'-------------------------------------------------------------------------------
Function RoundString(sIncoming, iDecPlaces)
	RoundString=Round(CDbl(Digits(sIncoming)),iDecPlaces)
End Function

'-------------------------------------------------------------------------------
' Function: RightOf
'
'  Returns all the string in the sSuperString which is to the right of the sSubString
'
' Parameters:
'  - sSuperString - the string to check inside of
'  - sSubString   - the string to search for
'
' Returns:
'  Either the complete sSupserString, if not found OR whatever is to the right 
'  of the sSubString
'
' Usage:
'  x = RightOf("123.456",".")
'  returns "456"
'-------------------------------------------------------------------------------
Function RightOf(sSuperString, sSubString)
	Dim sResult, iIndex
	sResult=sSuperString
	iIndex=InStr(1,sSuperString,sSubString)
	If iIndex > 0 Then
		sResult=Mid(sSuperString,iIndex+Len(sSubString))
	End If
	RightOf=sResult
End Function


'-------------------------------------------------------------------------------
' Function: LeftOf
'
'  Returns all the string in the sSuperString which is to the left of the sSubString
'
' Parameters:
'  - sSuperString - the string to check inside of
'  - sSubString   - the string to search for
'
' Returns:
'  Either the complete sSupserString, if not found OR whatever is to the left
'  of the sSubString
'
' Usage:
'  x = LeftOf("123.456",".")
'  returns "123"
'-------------------------------------------------------------------------------
Function LeftOf(sSuperString, sSubString)
	Dim sResult, iIndex
	sResult=sSuperString
	iIndex = InStr(1,sSuperString,sSubString)
	If iIndex > 0 Then
		sResult=Mid(sSuperString,1,iIndex-1)
	End If
	LeftOf=sResult
End Function

'-------------------------------------------------------------------------------
' Function: PadLeft
'
'  Pads the left side of a string with a specified character, and sets the string
'  to a specific length
'
' Parameters:
'   - incomingString - the starting string
'   - padChar        - the character to pad the string with
'   - totalWidth     - the width for the final result
'
' Returns:
'  a string with the pad char added to the left, and then reduced to the target length
'
' Usage:
'   print PadLeft("123","0",6)
'   would print: 000123
'
'-------------------------------------------------------------------------------
Function PadLeft(incomingString,padChar,totalWidth)
   Dim result
   result = String(totalWidth,padChar)
   result = result & incomingString 
   result = Right(result,totalWidth)
   PadLeft = result
End Function

'-------------------------------------------------------------------------------
' Function: PadRight
'
'  Pads the right side of a string with a specified character, and sets the string
'  to a specific length
'
' Parameters:
'   - incomingString - the starting string
'   - padChar        - the character to pad the string with
'   - totalWidth     - the width for the final result
'
' Returns:
'  a string with the pad char added to the right, and then reduced to the target length
'
' Usage:
'   print PadRight("123","0",6)
'   would print: 123000
'
'-------------------------------------------------------------------------------
Function PadRight(incomingString,padChar,totalWidth)
   Dim result
   result = String(totalWidth, padChar)
   result = incomingString & result
   result = Left(result,totalWidth)
   PadRight = result
End Function

'-------------------------------------------------------------------------------
' Function: MakeTextCompareString
'
'  Shortcut to return a lower case string containing only letters and numbers
'  used in the List object to do fuzzy compares
'
' Parameters:
'   - stringToTransform - the string to transform
'
' Returns:
'  string of lower case letters and numbers from the stringToTransform
'
'-------------------------------------------------------------------------------
Function MakeTextCompareString(stringToTransform)
	Dim CharacterIndex, Result, CurrentChar
	stringToTransform=LCase(stringToTransform)
	Result=""
	For CharacterIndex=1 to Len(stringToTransform)
		CurrentChar=Mid(stringToTransform,CharacterIndex,1)
		If (CurrentChar=>"a" AND CurrentChar <="z") OR (CurrentChar=>"0" AND CurrentChar <="9") Then
			Result=Result&CurrentChar
		End If
	Next
	MakeTextCompareString=Result
End Function

'-------------------------------------------------------------------------------
' Function: MakeItemPrintable
'
'  Used for debugging. Attempts to render the data as a string. For instance, 
'  renders the array (1,2,3) as "Array(1,2,3)"
'
' Parameters:
'   - theItem - the item to render into a printable form
'
' Returns:
'  a best guess string representation of the item
'
'-------------------------------------------------------------------------------
Function MakeItemPrintable(theItem)
	Dim result, looper, objectType
	
	If IsArray(theItem) Then
		For looper = 0 to UBound(theItem)
			result = result & ", " & MakeItemPrintable(theItem(looper))
		Next
		result = "Array(" & Mid(result,3) & ")"
	ElseIf IsNull(theItem) Then
		result = "*NULL*"
	ElseIf IsEmpty(theItem) Then
		result = "*EMPTY*"
	ElseIf IsObject(theItem) Then
		objectType = GetClass(theItem)
		If objectTYpe <> "" Then		' inside this IF is where we'd add processing for any other object types, eg Lists
			On Error Resume Next
			If NOT IsEmpty(theItem.ToString) Then
				result = theItem.ToString
			Else
				result = "Object of type " & objectType
			End If
			On Error Goto 0
		Else
			result = "Object of unknown type"
		End If
	Else
		result = theItem
	End If
	MakeItemPrintable = result
End Function

'-------------------------------------------------------------------------------
' Function: MakeItemPrintableWithoutExtrernalArray
'
'  Used for debugging - see MakeItemPrintable
'  I created this because of the space limitations in the debugging window in
'  QTP. Removing the "Arrray(" at the beginning let me see more of the data.
'
' Parameters:
'   - theArray - Can actually be ANY kind of data. 
'
' Returns:
'  The string of the data sans any external "Array(" wrapper
'
'-------------------------------------------------------------------------------
Function MakeItemPrintableWithoutExtrernalArray(theArray)
	Dim result
	result = MakeItemPrintable(theArray)
	If Mid(result,1,6)="Array(" Then
		result = Mid(result,7,Len(result)-7)
	End If
	MakeItemPrintableWithoutExtrernalArray = result
End Function





'***********************************************************
' aggregating q:\utils\libs\extensionsfiles.vbs

'###############################################################################
' Library: ExtensionsFiles.vbs
'
' About: Basic extensions to VBScript/QTP - Akien MacIain
'  Things that *should* have been in VBScript gathered together
'  Functions supporting working with files
'
'  Copyright (C) 2008, 2009, 2010 Akien MacIain
'
'  This program is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'###############################################################################
'Option Explicit

'-------------------------------------------------------------------------------
' Function: FrameworkDetectExtensionsFiles
'   Utility function for the Framework Compilation checking utility
'
' Returns:
'  (integer) always returns 1
'-------------------------------------------------------------------------------
Function FrameworkDetectExtensionsFiles()
	FrameworkDetectExtensionsFiles = 1
End Function

'===============================================================================
' Section: Constants and Globals

'===============================================================================
'-------------------------------------------------------------------------------
' Constants: Public Constants
'  Constants published by the library
'
'   File System Object Constants:
'   FSO_OVERWRITE_ON_COPY        - Overwrite files or folders when copying
'   FSO_DONT_OVERWRITE_ON_COPY   - Do not overwrite files or folders when copying
'   FSO_OVERWRITE_ON_CREATE      - Overwrite files or folders on create
'   FSO_DONT_OVERWRITE_ON_CREATE - Do not overwrite files or folders on create
'   FSO_FORCE_DELETE             - Force delete when the read-only attribute is set
'   FSO_DONT_FORCE_DELETE        - Do no delete when the read only attribute is set
'   FSO_IOMODE_FOR_READING       - Open files for reading only
'   FSO_IOMODE_FOR_WRITING       - Open files for writing
'   FSO_IOMODE_FOR_APPENDING     - Open files for writing and append
'   FSO_IO_CAN_CREATE_NEW_FILE   - Create a new text file if it doesn't already exist
'   FSO_IO_CANT_CREATE_NEW_FILE  - If the specified file doesn't exist, don't create it
'   FSO_IO_FORMAT_SYSTEM_DEFAULT - Open file using system default format
'   FSO_IO_FORMAT_UNICODE        - Open file as Unicode
'   FSO_IO_FORMAT_ASCII          - Open file as ASCII
'-------------------------------------------------------------------------------
Public Const FSO_OVERWRITE_ON_COPY        = True
Public Const FSO_DONT_OVERWRITE_ON_COPY   = False
Public Const FSO_OVERWRITE_ON_CREATE      = True
Public Const FSO_DONT_OVERWRITE_ON_CREATE = False
Public Const FSO_FORCE_DELETE             = True
Public Const FSO_DONT_FORCE_DELETE        = False
Public Const FSO_IOMODE_FOR_READING       = 1
Public Const FSO_IOMODE_FOR_WRITING       = 2
Public Const FSO_IOMODE_FOR_APPENDING     = 8
Public Const FSO_IO_CAN_CREATE_NEW_FILE   = True
Public Const FSO_IO_CANT_CREATE_NEW_FILE  = False
Public Const FSO_IO_FORMAT_SYSTEM_DEFAULT = -2
Public Const FSO_IO_FORMAT_UNICODE        = -1
Public Const FSO_IO_FORMAT_ASCII          = 0

'===============================================================================
' Section: Public Functions
'   Functions that are exported by the library
'===============================================================================

'-------------------------------------------------------------------------------
' Function: GetMD5ForFile 
'   Compute the MD5 hash for a given file
'
'  The function makes use of the DotNetFactory to give us access to the .NET
'  MD5 hashing algorithms.  This is many orders of magnitude faster then a
'  native VBScript MD5 algorithm.
'
' Parameters:
'  - fileName: (String) The name of the file you want to process.  NOTE that the
'               function assumes that the file exists.
'
' Returns:
'  (String) The MD5 hash for the file
'-------------------------------------------------------------------------------
Public Function GetMD5ForFile (fileName)
	Dim FileIO
	Dim MD5Provider
	Dim fileBytes
	Dim hashBytes
	Dim i
	Dim hash
   Dim d

'   MercuryTimers.Timer("bft").Start
	Set FileIO = DotNetFactory.CreateInstance("System.IO.File")
	Set MD5Provider = DotNetFactory.CreateInstance("System.Security.Cryptography.MD5CryptoServiceProvider")
	
	Set fileBytes = FileIO.ReadAllBytes(fileName)
	Set hashBytes = MD5Provider.ComputeHash(fileBytes)

	hash = ""
	For i = 0 to hashBytes.Length - 1
		hash = hash & hashBytes.GetValue(CInt(i)).ToString("x2")
	Next

	GetMD5ForFile = hash
'	d = MercuryTimers.Timer("bft").Stop
'   print "MD5 for file " & filename & " took " & d & "ms"

	Set FileIO = Nothing
	Set MD5Provider = Nothing
	Set fileBytes = Nothing
	Set hashBytes = Nothing
End Function

'-------------------------------------------------------------------------------
' Function: ValidateManifestFile
'   Validates that all the files listed in a manifest have the correct file
'   version
'
'   It does this by comparing the MD5 hash for the file to the hash stored in the
'   manifest. The format for the manifest is <file_path>,<md5>, one entry per line
'
' Parameters:
'  - fileName: (String) The name of the manifest file you want to process.
'
' Returns:
'  (Boolean) True if the file validates, false if it does not
'-------------------------------------------------------------------------------
Public Function ValidateManifestFile (fileName)
   Dim i
   Dim checkFailed
   Dim manifestFile
   Dim manifest

   checkFailed = False
   Set manifestFile = FSO().OpenTextFile(fileName, 1)

   Do While manifestFile.AtEndOfStream <> True
      manifest = Split(manifestFile.ReadLine, ",")
      If GetMD5ForFile(manifest(0)) <> manifest(1) Then
         checkFailed = True
         Exit Do
      End If
   Loop

   manifestFile.Close
   Set manifestFile = Nothing
 
   If checkFailed Then
      ValidateManifestFile = False
   Else
      ValidateManifestFile = True
   End If
   
End Function

'-------------------------------------------------------------------------------
' Function: FileNamePortionExtractFromPath
'  
'  takes C:\foo\bar\babble.txt, .txt returns babble
'
' Parameters:
'   - stringFileName - string - full path
'   - extension - string - the extension expected at the end
'
' Returns:
'   - string with just file name part (assumes Windows directory seperation characters)
'
'-------------------------------------------------------------------------------
Function FileNamePortionExtractFromPath(stringFileName, extension)
   Dim c, d
   
   c = LeftOf(stringFileName, extension)
   d = c
   While Contains(d,"\")
      d = RightOf(d,"\")
   Wend

   FileNamePortionExtractFromPath = d
End Function

'-------------------------------------------------------------------------------
' Function: FileNameDataExtractor
'  
'  takes C:\foo\bar\babble.txt, returns a dict with all file data (see below)
'
' Parameters:
'   - stringFileName - string - full path
'
' Returns: a dict with these fields:
'     - Key:passedFileSpec
'     - Key:fullName
'     - Key:drive
'     - Key:path
'     - Key:name
'     - Key:extension
'     - Key:exists
'     - Key:fullPath     
'     - Key:arrayOfDirectories
'     - Key:fso.fileObject
'     - Key:fso.parentFolderObject
'     - Key:attributes
'     - Key:dateCreated
'     - Key:dateLastAccessed
'     - Key:dateLastModified
'     - Key:size
'     - Key:type
'
'-------------------------------------------------------------------------------
Function FileNameDataExtractor(byVal stringFileName)

   Dim resultDir, temp, i
   Set resultDir = CreateObject("Scripting.Dictionary")
   
   resultDir("passedFileSpec") = stringFileName
   resultDir("fullPath") = FSO().GetAbsolutePathName(stringFileName)
   resultDir("exists") = FSO().FileExists(stringFileName)

   If resultDir("exists") Then
      ' fetch everything from the disk
      Dim theFile, theDir
      stringFileName = resultDir("fullPath")   
      resultDir("fullName") = FSO().GetFileName(stringFileName)

      Set resultDir("fso.fileObject") = FSO().GetFile(stringFileName)
      Set resultDir("fso.parentFolderObject") = FSO().GetFile(stringFileName).ParentFolder
      Set theFile = resultDir("fso.fileObject")
      Set theDir  = resultDir("fso.parentFolderObject")

      resultDir("name") = FSO().GetBaseName(stringFileName)
      resultDir("extension") = FSO().GetExtensionName(stringFileName)

      resultDir("drive") = Left(theFile.Drive,1)
      
      resultDir("path") = theDir.Path & "\"
      If Mid(resultDir("path"),2,1)=":" Then
         resultDir("path") = Mid(resultDir("path"),3)
      End If
      temp = resultDir("path")
      If Left(temp,1) = "\" Then
         temp = Mid(temp,2)
      End If
      If Right(temp,1) = "\" Then
         temp = Mid(temp,1,Len(temp)-1)
      End If
      resultDir("arrayOfDirectories") = Split(temp,"\") 
      
      resultDir("attributes") = theFile.Attributes
      resultDir("dateCreated") = theFile.DateCreated 
      resultDir("dateLastAccessed") = theFile.DateLastAccessed
      resultDir("dateLastModified") = theFile.DateLastModified
      resultDir("size") = theFile.Size
      resultDir("type") = theFile.Type

   Else
      ' apperently, we're on our own. the file doesn't actually exist 
      ' on the disk so we only have the file name to go on
   End If
   
End Function



'-------------------------------------------------------------------------------
' Function: FSO()
'  
'  Returns a file system object. Uses the GlobalDictionary to store it so it 
'  doesn't have to recreate it every time
'
' Parameters:
'   - None
'
' Returns:
'  The FSO
'
' Notes:
'  Assumes a GlobalDictionary to store it in
'-------------------------------------------------------------------------------
Function FSO()
	If IsEmpty(GlobalDictionary("fso")) Then
		GlobalDictionaryAdd "fso",CreateObject("Scripting.FileSystemObject")
	End If
	Set FSO = GlobalDictionary("fso")
End Function

'-------------------------------------------------------------------------------
' Sub: WriteArray
'
'  Writes an array of strings to a file name
'
' Parameters:
'   - arrayToWrite - the array of strings to write
'   - fileName     - the file to write to
'
' Notes:
'  Dependency on the List object elsewhere in this library set
'-------------------------------------------------------------------------------

Sub WriteArray (arrayToWrite, fileName)
	Dim writeList
	Set writeList = NewList()
	If IsObject(arrayToWrite) Then
		writeList.l = arrayToWrite.l
	Else
		writeList.l = arrayToWrite
	End If	
	writeList.WriteToFile fileName
	writeList = Empty
End Sub

'-------------------------------------------------------------------------------
' Sub: MakeFolder
'
'  Shortcut that makes the named folder (I got tired of always creating an FSO)
'
' Parameters:
'   - folderName - The name of the folder to create
'
' Notes:
'  Dependency on the FSO() function elsewhere in this library set
'-------------------------------------------------------------------------------
Sub MakeFolder(folderName)
	'Dim fso
	'Set fso = CreateObject("Scripting.FileSystemObject")
	If FSO().FolderExists(folderName) = False Then
		FSO().CreateFolder(folderName)
	End If
End Sub

'-------------------------------------------------------------------------------
' Sub: Run
'
'  Runs a command line
'
' Parameters:
'   - commandLine - What to run
'
' Notes:
'  Dependency on WSSHell() defined elsewhere in this library set
'-------------------------------------------------------------------------------
Sub Run (commandLine)
	WSSHell().Run commandLine,1,False
End Sub

'-------------------------------------------------------------------------------
' Sub: RunAndWait
'
'  Takes a command line, runs it, and waits for it to complete
'
' Parameters:
'   - commandLine - the command line to execute
'
' Notes:
'  Dependency on WSShell() defined elsewhere in this library set
'-------------------------------------------------------------------------------
Sub RunAndWait (commandLine)
	WSSHell().Run commandLine,1,True
End Sub

'-------------------------------------------------------------------------------
' Function: FileExists
'
'  Another FSO related short cut. This one returns file existance
'
' Parameters:
'   - sFilespec - the file to check the existance of
'
' Returns:
'  True or False based on whether the file could be found
'
' Notes:
'  This current incarnation depends on FSO() defined elsewhere in this library set
'-------------------------------------------------------------------------------
Function FileExists(sFilespec)
	'Dim fso
	'Set fso = CreateObject("Scripting.FileSystemObject")
	FileExists = FSO().FileExists(sFilespec)
End Function

'-------------------------------------------------------------------------------
' Sub: VerifyFileExist
'
'  Logs a fatal error if the specified file does not exist
'
' Parameters:
'   - fileName
'
' Notes:
'  Depends on FileExists defined elsewhere in this library 
'-------------------------------------------------------------------------------
Sub VerifyFileExist(fileName)
	If not FileExists(fileName) Then
        LogFatal "routine=>VerifyFileExist|message=>File does not exist: " & fileName
	End If
End Sub

'-------------------------------------------------------------------------------
' Sub: FileDelete
'
'  Shortcut to delete a file
'
' Parameters:
'   - filespec - the file to delete
'
' Notes:
'  Another shortcut. this one depends on FSO() defined elsewhere in this library set
'-------------------------------------------------------------------------------
Sub FileDelete(filespec)
	'Dim fso
	'Set fso = CreateObject("Scripting.FileSystemObject")
	If FileExists(sFileSpec) Then
		FSO().DeleteFile(filespec)
	End If
End Sub

'-------------------------------------------------------------------------------
' Function: WriteTestImage [DEPRICATED]
'
'  Writes an image of the desktop to a file in the framework logging directory
'  fileName = FetchLogDir() & ReturnYYYYMMDDHHMM & "." & TestCase("name") & ".StoredImage.png" 
'
' Parameters:
'   - None
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function WriteTestImage ()
	Dim fileName
	fileName = FetchLogDir() & ReturnYYYYMMDDHHMM & "." & TestCase("name") & ".StoredImage.png" 
	Desktop.CaptureBitmap fileName, True
End Function

'-------------------------------------------------------------------------------
' Function: WriteTestDataFile
'
'  Writes an array of data to a file in the framework logging directory
' fileToWriteTo = FetchLogDir() & ReturnYYYYMMDDHHMM & "." & TestCase("name") & ".FoundData.txt"
'
' Parameters:
'   - arrayToWrite
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function WriteTestDataFile(arrayToWrite)
	Dim fileToWriteTo
	fileToWriteTo = FetchLogDir() & ReturnYYYYMMDDHHMM & "." & TestCase("name") & ".FoundData.txt"
	WriteArray arrayToWrite, fileToWriteTo
	LogDebug "routine=>WriteTestDataFile;message=>" & "File:[" & fileToWriteTo ' & "] Data written:[" & arrayToWrite & "]" 
End Function














'***********************************************************
' aggregating q:\utils\libs\extensionsdates.vbs

'###############################################################################
' Library: ExtensionsDates.vbs
'
' About: Basic extensions to VBScript/QTP - Akien MacIain
'  Things that *should* have been in VBScript gathered together
'  Functions supporting working with dates
'
'  Copyright (C) 2008, 2009, 2010 Akien MacIain
'
'  This program is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'###############################################################################
'Option Explicit

'-------------------------------------------------------------------------------
' Function: FrameworkDetectExtensionsDates
'   Utility function for the Framework Compilation checking utility
'
' Returns:
'  (integer) always returns 1
'-------------------------------------------------------------------------------
Function FrameworkDetectExtensionsDates()
	FrameworkDetectExtensionsDates = 1
End Function


'===============================================================================
' Section: Public Functions
'   Functions that are exported by the library
'===============================================================================
' 1) Assignments
' 2) Typing
' 3) Conversion


'-------------------------------------------------------------------------------
' Function: returnyyyymmddhhmm
' 
'  strips off date time NOW in the form yyyymmddhhmmss
'
' Parameters: 
'	- none
'   
' Returns: 
'	'yymmddhhmmss'
'
' Notes:
'   Assumes DateTime to be 'yyyymmddhhmmss' (12 digits)
'-------------------------------------------------------------------------------
Function returnyyyymmddhhmm() 'return 14 digits(yymmddhhmmssxx), last 2-digits is a random number
   returnyyyymmddhhmm = right(year(Date), 2)
   returnyyyymmddhhmm = returnyyyymmddhhmm & PrefixWithZeros(Month(Date),2)
   returnyyyymmddhhmm = returnyyyymmddhhmm & PrefixWithZeros(Day(Date),2) 
   returnyyyymmddhhmm = returnyyyymmddhhmm & PrefixWithZeros(Hour(Time),2)
   returnyyyymmddhhmm = returnyyyymmddhhmm & PrefixWithZeros(Minute(Time),2)
   returnyyyymmddhhmm = returnyyyymmddhhmm & PrefixWithZeros(Second(Time),2)
	Randomize
   returnyyyymmddhhmm = returnyyyymmddhhmm & PrefixWithZeros(int(rnd*100),2)

'  (OR in 1 sentence)
'  returnyyyymmddhhmm = year(Date) & PrefixWithZeros(Month(Date),2) & PrefixWithZeros(Day(Date),2) & PrefixWithZeros(Hour(Time),2) & PrefixWithZeros(Minute(Time),2)
End Function

'-------------------------------------------------------------------------------
' Function: StripOffDateTimeSuffix
' 
'  strips off date time suffix IF IT is suffixed and returns
'  If there is no such suffix, the whole value is returned
'
' Parameters: 
'	- 'DateTime suffix to be 'yyyyddmmhhmm' (12 digits)
'   
' Returns: 
'	Value stripped of date time suffix
'
' Notes:
'   Assumes DateTime suffix to be 'yyyyddmmhhmm' (12 digits)
'-------------------------------------------------------------------------------
Function StripOffDateTimeSuffix(lastName)
    Dim dateStr
    StripOffDateTimeSuffix = lastName
    If len(lastName) > 12 Then
        dateStr = right(lastName, 12)
        If (isNumeric(dateStr)) Then
            StripOffDateTimeSuffix = mid(lastName, 1, (len(lastName) - 12))
        End If
    End If
End Function

'-------------------------------------------------------------------------------
' Function: PrefixWithZeros
'
'	crude routien to prefix any numeric value  < 10 with 0
'
' Parameters: 
'	- none
'   
' Returns: 
'	Prefixed value
'
' Notes:
'   Lots of work needed if this is to be used as generic function
'-------------------------------------------------------------------------------
Function PrefixWithZeros(num, places)
   PrefixWithZeros = num
   If num < 10 Then
        PrefixWithZeros = "0" & num
   End If
End Function

'-------------------------------------------------------------------------------
' Function: GetDateString
'
' Gets the current date and time in the format YYYYMMDD-HHMI
'
' Parameters:
'   - none
'
' Returns:
'  - string - containing the current date in time as YYYYMMDD-HHMI
'-------------------------------------------------------------------------------
Public Function GetDateString
   Dim formatter
   Dim timestamp
   
   timestamp = Now()
   Set formatter = NewcvDateFormat
   GetDateString = formatter.FormatDate(timestamp, "YYYYMMDD") & "-" & formatter.FormatTime(timestamp, "HHMM")
End Function

'-------------------------------------------------------------------------------
' Function: PadExcelDateString
'
' Given a string taken from an excel spreadsheet and pads out the month and day
' if necessary to ensure the result is in the format MMDDYYY
'
' Parameters:
'   - excelDate - The date string as read out of an excel spreadsheet
'
' Returns:
'  - string - The date padded out to eight characters
'
'-------------------------------------------------------------------------------
Public Function PadExcelDateString (excelDate)
   Dim dateString
   Dim formatter

   If (Trim(excelDate) = "") Then
      PadExcelDateString = ""
      Exit Function
   End If

   excelDate = StrictDigits(excelDate)
   
   If Len(excelDate) = 8 and IsNumbersOnly(excelDate) Then
      PadExcelDateString = excelDate
      Exit Function
   End If
   
   Set formatter = NewcvDateFormat
   PadExcelDateString = formatter.FormatDate(excelDate, "MMDDYYYY")
End Function

'-------------------------------------------------------------------------------
' Function: PrettyPrintTimer 
'
'   This function takes in a time interval in milliseconds and returns a pretty
'   string in the format 'X Hour(s) Y Minute(s) Z second(s) Q millisecond(s)'
'
'   The function was originaly developed by Ryan Trudelle-Schwarz for www.mamanze.com
'
' Parameters:
'   - delta: (Integer) The interval in milliseconds to format
'
' Returns:
'  (String) A pretty printed string representing the time interval
'
'-------------------------------------------------------------------------------
Function PrettyPrintTimer(byVal delta)
   Dim intMilliSecond, intSecond, intMinute, intHour
   Dim strReturn

   strReturn =""  

   ' Determine the number of milliseconds.
   intMilliSecond = delta mod 1000

   ' Determine the number of seconds. This is not the second value
   ' yet, just the number of seconds.
   intSecond = Int(delta/1000)

   ' Determine the number of minutes, simply divide the total number
   ' of seconds by 60 and get the real number result.
   intMinute = Int(intSecond / 60)

   ' Now we modulus the seconds by 60 to form the seconds value.
   intSecond = intSecond mod 60

   ' Compute the Hours value by dividing the minutes by 60.
   intHour = Int(intMinute / 60)

   ' Compute the actual minute value by getting the modulus of the
   ' total number of minutes and 60.
   intMinute = intMinute mod 60

   ' If the timer took more then a hour then display the hours.
   If intHour > 0 Then
      If intHour = 1 Then
         strReturn = strReturn & intHour &" Hour "  
      Else
         strReturn = strReturn & intHour &" Hours "  
      End If
   End If

   ' If the timer took more then a minute then display the minutes.
   If intMinute > 0 Then
      If intMinute = 1 Then
         strReturn = strReturn & intMinute &" Minute "  
      Else
         strReturn = strReturn & intMinute &" Minutes "  
      End If
   End If

   ' If the timer took more then a second then display the seconds.
   If intSecond > 0 Then
      If intSecond = 1 Then
         strReturn = strReturn & intSecond &" Second "  
      Else
         strReturn = strReturn & intSecond &" Seconds "  
      End If
   End If

   ' If the timer took more then a millisecond then display the
   ' milliseconds. Also, if the script took no time then display 0
   ' milliseconds.

   If strReturn ="" OR intMilliSecond > 0 Then
      If intMilliSecond = 1 Then
         strReturn = strReturn & intMilliSecond &" MilliSecond"  
      Else
         strReturn = strReturn & intMilliSecond &" MilliSeconds"  
      End If
   End If

   PrettyPrintTimer = strReturn
End Function

'-------------------------------------------------------------------------------
' Function: GetYesterdaysDate
' 
' Get yesterday's date using the format specified by the format parameter
'
' Parameters:
'   - Format - a cvDateFormat compliant format string
'
' Returns:
'  - string - containing yesterdays date in the format specified by format
'-------------------------------------------------------------------------------
Function GetYesterdaysDate(format)
   Dim formatter
   Set formatter = NewcvDateFormat
   
   GetYesterdaysDate = formatter.FormatDate(Date - 1, format)
End Function
'-------------------------------------------------------------------------------
' Function: targetdate
' 
' Get older date using the numberofyears parameter
'
' Parameters:
'   - numberofyears 
'
' Returns:
'  - integer - containing older date.
'-------------------------------------------------------------------------------
Function targetdate (numberofyears)
Dim currentdate,datearray,targetyear
    currentdate = date 
        If numberofyears = " " Then
	       numberofyears = 0
        End If

    datearray = split (currentdate, "/")
    targetyear = datearray(2) - numberofyears
    targetdate = datearray (0) & "/" & datearray(1) & "/" & targetyear

End Function
'-------------------------------------------------------------------------------
' Function: MMDDYYYY
'  Retuns the date formatted as MMDDYYYY
'
' Parameters: 
'  - aDate - a date, e.g. 7/9/2008
'
' Returns:
'  return date in format: "MMDDYYYY"
'-------------------------------------------------------------------------------
Function MMDDYYYY(aDate) ' return date in format: "MMDDYYYY"
   Dim aday, amonth
 aday = day(aDate)

 if aday <10 then
	aday = "0" & aday
 end if
 amonth = month(aDate)
 if amonth <10 then
	amonth = "0" & amonth 
 end if

 MMDDYYYY = amonth & aday & year(aDate) 
End Function
'-------------------------------------------------------------------------------
' Function: MMDDYYYY_slash(aDate)
' parameter : a date, e.g. 7/9/2008 or date Function
' Returns:
'  	return date in format: "MM/DD/YYYY"
'-------------------------------------------------------------------------------
Function MMDDYYYY_slash(aDate) ' return date in format: "MM/DD/YYYY"
 Dim aday, amonth
 aday = day(aDate)

 if aday <10 then
	aday = "0" & aday
 end if
 amonth = month(aDate)
 if amonth <10 then
	amonth = "0" & amonth 
 end if

 MMDDYYYY_slash = amonth & "/" & aday & "/" & year(aDate) 
End Function
'-------------------------------------------------------------------------------
' Function: DateDiffMMDDYYYY
'  Get the difference between two dates formatted as MMDDYYYY
'
' Parameters: 
'  date1 - date to subtract from
'  date2 - the date to subtract
'
' Returns: 
'  0 if same;  =1 if date1 > date2;  =-1 if date1 < data2
'-------------------------------------------------------------------------------
function DateDiffMMDDYYYY(date1, date2)
	date1 = right(date1,4) & left(date1, 4) 
	date2 = right(date2,4) & left(date2, 4) 
	'if right(date1,4) = right(date2, 4) then
	if date1 = date2 then
		DateDiffMMDDYYYY = 0
	elseif date1 > date2 then
		DateDiffMMDDYYYY = 1
	else
		DateDiffMMDDYYYY = -1
	end if
end function

'-------------------------------------------------------------------------------
' Function: DateAddDaysMMDDYYYY
'
'  similar to DateAdd, except date format is MMDDYYYY
'
' Parameters:
'   - dateIn  - the date to add days to
'   - number  - the number of days to add
'
' Returns:
'  a date in format MMDDYYYY
'
' Usage:
'  DateAddDaysMMDDYYYY("02132007", -1)  would return "02132006"
'-------------------------------------------------------------------------------
function DateAddDaysMMDDYYYY(dateIn, number)
	dim yyyy, mm, dd, date2
	yyyy = right(dateIn, 4)
	mm   = left(dateIn, 2)
	dd   = mid(dateIn, 3, 2)
	date2 = mm & "/" & dd & "/" & yyyy

	DateAddDaysMMDDYYYY = MMDDYYYY( dateAdd("d", number, date2) )
end function




'***********************************************************
' aggregating q:\utils\libs\vbscript++.vbs

'###############################################################################
' Library: VBScript++.vbs
'
' THIS IS A KEY ARCHITECTURAL COMPONENT, THE ARCHITECT SHOULD BE NOTIFIED OF 
' CHANGES TO THIS FILE. 
'
' About: 
'  True object support for VBScript
'  Copyright (C) 2008, 2009, 2010 Akien MacIain
'
'  This program is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' Feature List:
'  - Implements Bag object, the substrate for all BagClasses and BagObjects
'  - Including the ability to define a class
'  - And to create an instance of a class
'  - To call methods and properties of a class
'  - To determine if an object is a Bag, BagClass or BagInstance
'  - To create from a dictionary or from a text string a BagDict
'
' Usage:
'
'###############################################################################
'Option Explicit

'-------------------------------------------------------------------------------
' Function: FrameworkDetectVBScriptPlusPlus
'   Utility function for the Framework Compilation checking utility
'
' Returns:
'  (integer) always returns 1
'-------------------------------------------------------------------------------
Function FrameworkDetectVBScriptPlusPlus()
	FrameworkDetectVBScriptPlusPlus = 1
End Function

'============================================================================================================
' OBJECT MODEL SPECIFIC EXTENSIONS - VBScript++.vbs

'-------------------------------------------------------------------------------
' Function: IsBag
'
'  Returns true if item is an object, and the object is a bag_dict
'
' Parameters:
'   - incoming - item to test
'
' Returns:
'     True/False
'
' Usage:
'
'  If IsBag(foo) Then
'     ...
'
'-------------------------------------------------------------------------------
Function IsBag(incoming)
   Dim result
   result = False
   If IsDict(incoming) Then
      If incoming.Item("self.is_bag_dict") = True Then
         result = True
      End If
   End If
   IsBag = result
End Function


'-------------------------------------------------------------------------------
' Function: BagDictMake
'
'  Makes and returns a bag dict. Not a bag class, but the container that makes
'  those work.
'
' Parameters:
'   - args - A DictMake compatible argument or dictionary
'
' Returns:
'  A Bag Dict object with whatever keys have been handed over in args copied to it
'  via a shallow copy
'
' Usage:
'  Set myDict = BagDictMake("my_string=>Hello world!")
'
'-------------------------------------------------------------------------------
Function BagDictMake (args, metaArgs)
   If IsBagDict(args) Then
      Set BagDictMake = args
   Else
      Set BagDictMake = New BagDict
      BagDictMake.ApplyKeys(args)
   End If
   If NOT IsNull(metaArgs) Then
      BagDictMake.ApplyMetadata(metaArgs)
   End If
End Function

'-------------------------------------------------------------------------------
' Sub: BagDictCreate
'
'  Like DictMake, alters the incoming args to be the bag dict
'
' Parameters:
'   - args - item to be rendered as a dict
'
' Usage:
'  foo = "thingie=>hello world"
'  DoIt foo
'  Sub DoIt(args)
'     BagDictCreate args
'     myThingie = args("thingie")
'
'-------------------------------------------------------------------------------
Sub BagDictCreate(args, metaArgs)
   If NOT IsBagDict(args) Then
      Set args = BagDictMake(args, metaArgs)
   End If
End Sub

'-------------------------------------------------------------------------------
' Function: IsBagDict
'
'  Returns true if the passed item is a bag dict
'
' Parameters:
'   - args - the item to be checked
'
' Returns:
'  True if is a bag dict, else False
'
' Usage:
'  If NOT IsBagDict(args) Then
'     ...
'
'-------------------------------------------------------------------------------
Function IsBagDict(args)
   IsBagDict = False
   If IsObject(args) Then
      IsBagDict = (args.Item("self.is_bag_dict") = True)
   End If
End Function

'-------------------------------------------------------------------------------
' Function: BagClassMake
'
'  Creates a bag class on top of a bag dict. Might be used to hack up a class
'  (Was used that way during debugging)
'
' Parameters:
'   - nameToUse - the name for the new class. Typically the same name as the variable
'                 the class is stored in. See usage information for additional information
'   - args      - DictMake compatible arguments to be added to the dictionary
'
' Returns:
'  A Bag Class object
'
' Usage:
'  Dim itemToMakeIntoClass
'  Set itemToMakeIntoClass = BagClassMake("classFile",null,"self.is_virtual=><eval>False") 
'
'-------------------------------------------------------------------------------
Function BagClassMake (nameToUse, args, metaArgs)
   Set BagClassMake = BagDictMake(NULL)
   Set BagClassMake = BagClassMake.MakeClass (nameToUse, args, metaArgs)
End Function

'============================================================================================================
' OBJECT MODEL DICTIONARY EXTENSION - VBScript++.vbs
'============================================================================================================

'-------------------------------------------------------------------------------
' Class: BagDict
' 
'  This class serves two purposes, and since VBScript doesn't allow inheritence, 
'  these two tightly related functions are implemented in one class. You can however
'  use one without the other. 
'
'  Use 1: Add functionality to the Dictionary object. Does things like check to see
'  if an item exists before calling Dictionary.Add that should have been built in.
'  A whole host of dictionary extensions are implemented this way. See below for specifics.
'  This is done by implementing a class which contains a dictionary, and then creating
'  "pass through" calls for each piece of Dictionary functionality. These calls are
'  modififed as needed to support the changed functionality (for instance, referencing
'  any key (even if it does not exist, such as in If IsEmpty(myDict("foo")) Then...)
'  causes the dictionary object to create that key. In this implementation, referencing
'  a non existant key does not create it. This also implements things like ItemIndex and
'  KeyIndex which allow you to reference an item in the item or key arrays by it's index
'  number. Detailed documentation is implemented below.
'
'  Use 2: Creates a new kind of "object". Since languages like C++ implement objects
'  with inheritence and polymorphism, and VBScript does not, and since C++ implements
'  it's objects as data structures with "hidden fields" (function pointers to
'  method calls and such), I decided to mash those ideas together and implement 
'  them in VBScript. Dictionary keys prefixed with "method." refer to method calls.
'  Property keys are prefixed by "prop." objectMetadata is prefixed with "self."
'
'  Properties and methods which are defined to go with the class are defined and 
'  implemented as follows
'     (start example)
'     ' this defines the container for the "new bag class"
'     Dim classFile 
'
'     ' this uses the classFileSystemItem bag class object to create a new bag class object
'     Set classFile = classFileSystemItem.MakeClass("classFile","self.is_virtual=><eval>False") 
'
'     ' then we apply a property to the class object:
'     classFile.ApplyProp "Exist","get","return_type=>Boolean"
'
'     ' and finally we define the code to be called. It is by default named
'     ' in the form: classname_propertyname_direction and ALWAYS takes 2 arguemnts: self and args
'     Function ClassFile_Exist_Get(self,args)
'        ClassFile_Exist_Get = ClassFileSystemItem_FSO.FileExists(self("file_name"))
'     End Function
'
'     ' creating an instance then looks like this:
'     Set myFile = classFile.NewObject ("file_name=>c:\foo.txt", NULL)
' 
'     ' calling a property then looks like this:
'     myResult = myFile.Prop ("Exist", NULL)
'     (end example)
'
'  The use of DictMake arguments (see Sub DictMake elsewhere) then allows us
'  to implement a simple form of polymorphism using conditionals within the 
'  called routine.
'
'  Items implemented using this approach are called Bag Classes or Bag Objects.
'  They're both BagDict objects, with different objectMetadata.
'
'  The dictionary passthrough code was stolen from Tarun Lalwani, published at:
'  http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/
'
' IMPORTANT:
'  Please note that all keys are lower case. Usually in the form component.word_word_word
'  component can refer to any grouping of attributes. In this example: self.foo_bar 
'  self is the component. For *variable names* using our coding standards, the 
'  local identifier within the component would be fooBar, but because of the all
'  lower case rule, foo_bar is used instead.
'
'  All of the classes and instances worked with using this code are intended to be 
'  dictionaries carrying data conforming to the bag_object model:
'
'     key:self.is_bag_object                 - required for all 
'
'     key:self.class_name                    - string: name of the class of self
'     key:self.inherits_from                 - bag_object class
'     key:self.inheritence_list              - dict: who else does this class inherit from? (grandparents)
'     key:self.is_instance                   - boolean: is this an instance?
'     key:self.is_virtual                    - boolean: is this a virutal class? (in the C++ sense)
'
'     key:method.<name>                      - string: function pointer to a method (subroutine, does not return a value)
'     key:prop.<name>                        - string: function pointer to a property (function, does return a value)
'     key:prop.<name>.return_type            - string: the return type (boolean, string, array, etc)
'     key:class.<name>                       - any type: items shared across members of a class (TBD: NOT YET IMPLEMENTED)
'     key:private.<name>                     - any type: private to this class (TBD: NOT YET IMPLEMENTED)
'     key:protected.<name>                   - any type: protected data (TBD: NOT YET IMPLEMENTED)
'
'     key:method.constructor                 - string: special purpose method, called during ObjCreate
'     key:method.destructor                  - string: special purpose method, called during ObjDestroy
'
'     key:self.is_hacked                     - boolean: special flag that indicates this dict didn't start life as a bag_object
'
'-------------------------------------------------------------------------------
Class BagDict

   '============================================================================================================
   Public objectMetadata
   Public localData
   Private localFSO

   '-------------------------------------------------------------------------------
   ' Method: Class_Initialize
   '
   '  Initializes instance. Sets up local data, sets modes on contained dicts. 
   '  Intialize event gets executed whenever a object is created
   '
   ' Parameters:
   '   - None
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  Called automatically, not called by programmer
   '
   '-------------------------------------------------------------------------------
   Sub Class_Initialize()
      Set localFSO = CreateObject("Scripting.FileSystemObject")
      
      Set objectMetadata = CreateObject("Scripting.Dictionary")
      Set localData = CreateObject("Scripting.Dictionary")
      
      objectMetadata.CompareMode = vbTextCompare
      localData.CompareMode = vbTextCompare

   End Sub

   '-------------------------------------------------------------------------------
   ' Method: Class_Terminate
   '
   '  Clears the data from the instance in preperation for destruction.
   '  Executed when the object is destroyed
   '
   ' Parameters:
   '   - None
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  Called automatically, not called by programmer
   '
   '-------------------------------------------------------------------------------
   Sub Class_Terminate()
      'call destructor
      If objectMetadata("self.is_instance") = True Then
         me.Method "Destroy", NULL
      End If
      
      'Remove all the keys
      objectMetadata.RemoveAll
      localData.RemoveAll
      
      'Destroy the dictionaries
      Set objectMetadata = Nothing
      Set localData = Nothing
   End Sub

   '============================================================================================================
   ' EXTENSIONS TO FUNCTIONALITY BY AKIEN WHICH I WISH ALL DICTS HAD
   '============================================================================================================

   '-------------------------------------------------------------------------------
   ' Property: ItemIndex
   '
   '  This gets around the inability to do dict.Items(n) (which should have been allowed
   '  but which the dictionary object tries to treat as a key reference rather than an 
   '  index into the Items array) 
   '
   ' Parameters:
   '   - indexIntoArray - the index into the array of items
   '   - value          - item passed into Let/Set
   '
   ' Returns:
   '   - (Get) Whatever is storied at that item, or Empty if not found
   '
   ' Exceptions:
   '   - Will trigger an out of range error if you try to access something past the 
   '     end of the array
   ' 
   ' Usage:
   '  n = myDict.ItemIndex(2)
   '
   ' Notes:
   '  Only interacts with local data, not metadata
   '
   '-------------------------------------------------------------------------------
   Public Property Get ItemIndex(indexIntoArray)
      Assign ItemIndex, localData.Item(me.KeyIndex(indexIntoArray))
   End Property

   Public Property Let ItemIndex(indexKey, Value)
      Dim keyToUse
      Assign keyToUse, me.KeyIndex(indexKey)
      Assign localData(keyToUse), Value
   End Property

   Public Property Set ItemIndex(indexKey, Value)
      Dim keyToUse
      Assign keyToUse, me.KeyIndex(indexKey)
      Assign localData(keyToUse), Value
   End Property

   '-------------------------------------------------------------------------------
   ' Property: KeyIndex
   '
   '  This gets around the inability to do dict.Keys(n) (which should have been allowed
   '  but which the dictionary object tries to treat as a key reference rather than an 
   '  index into the key array) 
   '
   ' Parameters:
   '   - indexIntoArray - the index into the array of items
   '
   ' Returns:
   '   - (Get) Whatever is storied at that item, or Empty if not found
   '
   ' Exceptions:
   '   - Will trigger an out of range error if you try to access something past the 
   '     end of the array
   ' 
   ' Usage:
   '  n = myDict.KeyIndex(2)
   '
   ' Notes:
   '  Only interacts with local data, not metadata
   '
   '-------------------------------------------------------------------------------
   Public Property Get KeyIndex(indexIntoKeyArray)
      Dim theKeys 
      theKeys = localData.Keys

      Dim resolvedIndex 
      resolvedIndex = ResolveIndex(indexIntoKeyArray)
      
      Assign KeyIndex, theKeys(resolvedIndex)
   End Property

   '-------------------------------------------------------------------------------
   ' Method: ResolveIndex (Private)
   '
   '  Allows you to specify indicies that are negative, in order to fetch last or
   '  items counted from the end. e.g. -1 = the last entry in the list, -2 = second to last
   '  and so on.
   '
   ' Parameters:
   '   - oldIndex - the number we start with, positives are returned unchanged
   '
   ' Returns:
   '   - the resolved index
   '
   ' Usage:
   '  newIndex = me.ResolveIndex(oldIndex)
   '
   '-------------------------------------------------------------------------------
   Private Function ResolveIndex(oldIndex)
      If oldIndex < 0 Then
         oldIndex = (localData.Count - 1) + (oldIndex + 1)
      End If
      ResolveIndex = oldIndex
   End Function

   
   '-------------------------------------------------------------------------------
   ' Property: SetIfUndefined
   '
   '  Will push the value into the dict if a value with that key does not already exist
   '
   ' Parameters:
   '   - keyToUse - the key (not the keyindex) to check for
   '   - value    - the value to set if the key does not already exist
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  myObject.SetIfUndefined("foo") = "bar"
   '
   ' Notes:
   '  Only interacts with local data, not metadata
   '
   '-------------------------------------------------------------------------------
   Public Property Let SetIfUndefined(keyToUse, Value)
      If NOT localData.Exists(keyToUse) Then
         If IsObject(Value) Then
            Set localData(keyToUse) = Value
         Else
            localData(keyToUse) = Value
         End If
      End If
   End Property
   Public Property Set SetIfUndefined(keyToUse, Value)
         If IsObject(Value) Then
            Set localData(keyToUse) = Value
         Else
            localData(keyToUse) = Value
         End If
   End Property

   '-------------------------------------------------------------------------------
   ' Method: MakeBag
   '
   '  Createes a new bag, and performs a shallow copy of the content of args onto it
   '
   ' Parameters:
   '   - args     - DictMake compatible arguments to be added to the resulting bag dict
   '   - metaArgs - DictMake compatible metadata arguments to be added to the resulting bag dict
   '                (may be null)
   '
   ' Returns:
   '   - a bag dict with the contents of args added to it
   '
   ' Usage:
   '  Set myNewDict = MakeBag("this_key_gets=>This Value", NULL)
   '
   '-------------------------------------------------------------------------------
   Function MakeBag (args, metaArgs)
      Set MakeBag = BagDictMake(args, metaArgs)
   End Function 
   
   '-------------------------------------------------------------------------------
   ' Sub: MakeIntoBagDict
   '
   '  Used to make a set of args into a bag dict. Very much a Bag Dict centric
   '  version of DictMake
   '
   ' Parameters (required):
   '   - args     - DictMake compatible arguments to be rendered as a bag dict
   '   - metaArgs - DictMake compatible metadata arguments to be added to the resulting bag dict
   '                (may be null)
   '
   ' Returns:
   '   - args as a new bag dict (unless the item is already a bag dict)
   '
   ' Usage:
   '  Sub foo(args)
   '     me.MakeBag(args, NULL)
   '     myItem = args("my_item")
   '
   ' Notes:
   '  This is intended to be used within the class as a Bag Dict version of DictMake
   '
   '-------------------------------------------------------------------------------
   Sub MakeIntoBagDict(args, metaArgs)
      If NOT IsBagDict(args) Then
         Set args = BagDictMake(args, metaArgs)
      Else
         If NOT IsNull(metaArgs) Then
            args.ApplyMetaData(metaArgs)
         End If
      End If
   End Sub
   
   '-------------------------------------------------------------------------------
   ' Method: Withdraw
   '
   '  Removes a key and returns it's value. Used where you might be passed multiple
   '  sub items you wish to take action with, then remove from the dict, usually because
   '  the last action will be to merge that dict into another dict.
   '
   ' Parameters:
   '   - keyToUse - the key to attept to fetch a value for
   '
   ' Returns:
   '   - Either Empty or whatever was found for that key
   '
   ' Usage:
   '  myData = myDict.Withdraw("foo")
   '  newDict.Merge(myDict)
   '
   ' Notes:
   '  Only interacts with local data, not metadata
   '
   '-------------------------------------------------------------------------------
   ' this returns the value and removes it from the dict... sort of "uses it up"
   Function Withdraw(keyToUse)
      Withdraw = Empty
      If localData.Exists (keyToUse) Then
         If IsObject(localData(keyToUse)) Then
            Set Withdraw = objectMetadata(keyToUse)
         Else
            Withdraw = localData(keyToUse)
         End If
         localData.Remove keyToUse
      End If
   End Function 

   '-------------------------------------------------------------------------------
   ' Method: GetKeyOrItem
   '
   '  Returns either the value stored in the specified key, or if that key does not
   '  exist, and key:item does, use the value stored in key:item. If neither exists
   '  return nothing.
   '
   ' Parameters:
   '   - keyToFetch - string, the name of the key to check for
   '
   ' Returns:
   '   - See description
   '
   ' Usage:
   '  Assign x, myDict.GetKeyOrItem("foo")
   '
   '-------------------------------------------------------------------------------
   Function GetKeyOrItem(keyToFetch)
      If me.Exist("item") Then
         assign GetKeyOrItem, me("item")
      End If
      If me.Exist(keyToFetch) Then
         assign GetKeyOrItem, me(keyToFetch)
      End If
   End Function

   '-------------------------------------------------------------------------------
   ' Method: HasKeys
   '
   '  Returns T/F based on whether a bag dict contains all the keys specified. This
   '  is used to pre check the existance of keys before the code starts trying to 
   '  use them.
   '
   ' Parameters:
   '   - arrayOfKeysToCheckFor - can be either an array or a single key to check for
   '
   ' Returns:
   '   - T/F based on whether ALL the keys were found
   '
   ' Usage:
   '  If NOT myDict.HasKeys("a","b","c") Then
   '     ' report error here
   '
   ' Notes:
   '  Checks localData AND objectMetadata
   '
   '-------------------------------------------------------------------------------
   Function HasKeys(arrayOfKeysToCheckFor)
      Dim result, i
      result = True

      If NOT IsArray(arrayOfKeysToCheckFor) Then
         arrayOfKeysToCheckFor = Array(arrayOfKeysToCheckFor)
      End If
      
      For i = 0 to UBound(arrayOfKeysToCheckFor)
         If (localData.Exists(arrayOfKeysToCheckFor(i)) OR objectMetadata.Exists(arrayOfKeysToCheckFor(i))) = False Then
            result = False
            Exit For
         End If
      Next
            
      HasKeys = result
   End Function
   
   '-------------------------------------------------------------------------------
   ' Method: VerifyHasKeys
   '
   '  If not all keys found, raises an error. Reports class, missing keys, and 
   '  a passed failure message. This allows us to have a single line to verify
   '  that a routine has the data it needs.
   '
   ' Parameters:
   '   - arrayOfKeysToCheckFor - can be either an array or a single key to check for
   '   - failuremessage        - message to post in the event of a failure
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  Sub Foo(args)
   '     me.VerifyHasKeys Array("a","b","c"), "Failed"
   '
   ' Notes:
   '  Resulting message: FATAL: bag object of type classBase found to be missing needed keys. " & failuremessage
   '
   '  Checks localdata and objectMetadata
   '
   '-------------------------------------------------------------------------------
   Sub VerifyHasKeys(arrayOfKeysToCheckFor, failuremessage)
      If NOT me.HasKeys(arrayOfKeysToCheckFor) Then
         Err.Raise 1,"", "TBD:FATAL: bag object of type " & objectMetadata("self.class_name") & " found to be missing needed keys. " & failuremessage
      End If
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: CopyTo
   '
   '  Copies keys to a target dictionary or bag dict
   '  shallow copies the keys/values to another dict
   '
   ' Parameters (required):
   '   - targetdict - the target dictionary/bag dict to copy to
   '
   ' Parameters (allowed):
   '   - args key:keys_to_copy - array with the list of keys to copy
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  myDict.CopyTo otherDict, NULL
   '  myDict.CopyTo otherDict, Array("keys_to_copy",Array("foo","bar"))
   '
   ' Notes:
   '  Performs a shallow copy. Object references will remain references to the
   '  same objects
   '
   '  Does NOT copy objectMetadata
   '
   '-------------------------------------------------------------------------------
   Sub CopyTo(byref targetDict, args)
      me.CopyKeys localData, targetDict, args
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: CopyFrom
   '
   '  Copies keys from another dictionary or bag dict to this one
   '  shallow copies the keys/values from another dict
   ' Parameters (required):
   '   - sourceDict - the dictionary to copy from
   '
   ' Parameters (allowed):
   '   - args key:keys_to_copy - array with the list of keys to copy
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  myDict.CopyFrom otherDict, NULL
   '  myDict.CopyFrom otherDict, Array("keys_to_copy",Array("foo","bar"))
   '
   ' Notes:
   '  Performs a shallow copy. Object references will remain references to the
   '  same objects
   '
   '  Does NOT copy objectMetadata
   '-------------------------------------------------------------------------------
   Sub CopyFrom (byref sourceDict, args)
      CopyKeys targetDict, localData, args
   End Sub
   
   '-------------------------------------------------------------------------------
   ' Method: CopyKeys
   '
   '  Called by the other copy code. shallow copies keys from one dict to another
   '
   ' Parameters:
   '   - fromDict - dictionary or bag dict to copy from
   '   - toDict   - dictionary or bag dict to copy to
   '   - args     - arguments that may effect the copy. As of this writing, the 
   '                only one that is used is args key:keys_to_copy
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  myDict.CopyKeys dictToCopyFrom, dictToCopyTo, NULL
   '  myDict.CopyKeys dictToCopyFrom, dictToCopyTo, Array("keys_to_copy",Array("foo","bar"))
   '  myDict.CopyKeys myDict.objectMetaData, newDict.objectMetaData, NULL <-- use this to copy metadata
   '  myDict.CopyKeys myDict.objectMetaData, newDict.objectMetaData, Array("keys_to_copy",Array("self.is_instance","self.is_nusance"))
   '
   ' Notes:
   '  Does NOT copy objectMetadata by default
   '
   '-------------------------------------------------------------------------------
   Sub CopyKeys(fromDict, toDict, args)
      If NOT IsNull(fromDict) AND NOT IsEmpty(fromDict) Then

         On Error Resume Next
         Set fromDict = fromDict.localData
         Set toDict = toDict.localData
         On Error Goto 0
         
         DictCopy fromDict, toDict, args
         
      End If
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: RenderAsArray
   '
   '  Returns the contents of this dictionary as an array in the DictMake form:
   '  Array("key1","value1","key2","value2")
   '
   ' Parameters:
   '   - None
   '
   ' Returns:
   '   - DictMake compatible array in the form: Array("key1","value1","key2","value2"...)
   '
   ' Usage:
   '  MyArray = myObject.RenderAsArray()
   '
   ' Notes:
   '  Does NOT include objectMetadata
   '
   '-------------------------------------------------------------------------------
   Function RenderAsArray() ' returns Array("key1","value1") type result
      Dim result
      ' BREAKPOINT HERE!
      Set result = BagDictMake(NULL,NULL)
      Dim i
      For i = 0 to me.Count-1
         result("K" & i) = me.KeyIndex(i)
         result("I" & i) = me.ItemIndex(i)
      Next
      RenderAsArray = result.Items
   End Function

   '-------------------------------------------------------------------------------
   ' Method: RenderAsString
   '
   '  Returns the contents of this dictionary as string in the DictMake form:
   '  "key1=>value1|key2=>value2"
   '
   ' Parameters:
   '   - None
   '
   ' Returns:
   '   - DictMake compatible string in the form "key1=>value1|key2=>value2"
   '
   ' Usage:
   '  print "The dictionary contains " & myDict.RenderAsString()
   '
   ' Notes:
   '  Does NOT include objectMetadata
   '
   '-------------------------------------------------------------------------------
   Function RenderAsString() ' returns "key1=>value1|key2=>value2" type result
      Dim theArray, result, i
      theArray = me.RenderAsArray()
      result = ""

      For i = 0 to UBound(theArray) Step 2
         result = result & CString(theArray(i)) & "=>" & CString(theArray(i+1)) & "|"
      Next

      If Len(result) > 0 Then
         result = Mid(result,1,Len(result)-1)
      End If
      
      RenderAsString = result
   End Function

   '-------------------------------------------------------------------------------
   ' Method: RenderAsArrayOfString
   '
   '  Returns the contents of this dictionary as an array of strings in the form:
   '  Array("key1=>value1","key2=>value2"...) - intended use is for debugging, for
   '  examining dictionary contents in real time. An object inspector, if you will.
   '
   ' Parameters:
   '   - None
   '
   ' Returns:
   '   - an array of string in the form Array("key1=>value1","key2=>value2"...)
   ' 
   ' Usage:
   '  Place in the variable/expression watch window: myDict.RenderAsArrayOfString()
   '
   ' Notes:
   '  THIS ROUTINE RETURNS BOTH OBJECT METADATA AND LOCAL DATA!! THIS IS NOT TYPICAL
   '  OF THE REST OF THIS LIBRARY, AND IS IMPLEMENTED THIS WAY BECAUSE THIS IS
   '  INTENDED AS A DEBUGGING TOOL
   '
   '  ITEM KEYS WILL BE PREFACED WITH M: or O: TO INDICATE WHETHER AN ITEM IS
   '  METADATA OR OBJECT DATA
   '
   '-------------------------------------------------------------------------------
   Function RenderCompleteObjectAsArrayOfString()
      'RenderAsArrayOfString = CArrayOfStringFromDict(objectMetadata)
      
      Dim result, i, theKeys, theValues
      Set result = CreateObject("Scripting.Dictionary")
      
      theKeys = objectMetadata.Keys
      theValues = objectMetadata.Items      
      For i = 0 to UBound(theKeys)
         result( "M:" & theKeys(i) & "=>" & CString(theValues(i)) ) = True 
      Next
   
      theKeys = localData.Keys
      theValues = localData.Items      
      For i = 0 to UBound(theKeys)
         result( "O:" & theKeys(i) & "=>" & CString(theValues(i)) ) = True 
      Next

      RenderCompleteObjectAsArrayOfString = result.Keys

   End Function
   
   '-------------------------------------------------------------------------------
   ' Method: Debug
   '
   '  Shortcut for RenderCompleteObjectAsArrayOfString, used in debugging
   '  The reason for the shortcut is typing the full name into the debugger was 
   '  slowing us down too much
   '
   '-------------------------------------------------------------------------------
   Function Debug()
      Debug = RenderCompleteObjectAsArrayOfString()
   End Function

   '-------------------------------------------------------------------------------
   ' Method: RenderCompleteObjectAsString
   '
   '  Returns the contents of this dictionary as string in the form:
   '  "key1=>value1|key2=>value2".  Intended use is for debugging, for
   '  examining dictionary contents in real time. An object inspector, if you will.
   '
   ' Parameters:
   '   - None
   '
   ' Returns:
   '   - a string in the form "key1=>value1|key2=>value2"
   ' 
   ' Usage:
   '  Place in the variable/expression watch window: myDict.RenderCompleteObjectAsString()
   '
   ' Notes:
   '  THIS ROUTINE RETURNS BOTH OBJECT METADATA AND LOCAL DATA!! THIS IS NOT TYPICAL
   '  OF THE REST OF THIS LIBRARY, AND IS IMPLEMENTED THIS WAY BECAUSE THIS IS
   '  INTENDED AS A DEBUGGING TOOL
   '
   '  ITEM KEYS WILL BE PREFACED WITH M: or O: TO INDICATE WHETHER AN ITEM IS
   '  METADATA OR OBJECT DATA
   '
   '-------------------------------------------------------------------------------
   Function RenderCompleteObjectAsString()

      Dim result, i, theKeys, theValues
      Set result = CreateObject("Scripting.Dictionary")
      
      theKeys = objectMetadata.Keys
      theValues = objectMetadata.Items
      
      For i = 0 to UBound(theKeys)
         result = result & "M:" & CString(theKeys(i)) & "=>" & CString(theValues(i)) & "|"
      Next
   
      theKeys = localData.Keys
      theValues = localData.Items  
      
      For i = 0 to UBound(theKeys)
         result = result & "O:" & CString(theKeys(i)) & "=>" & CString(theValues(i)) & "|"
      Next

      RenderCompleteObjectAsString = Mid(result,1,Len(result)-1)
      
   End Function
   
   '============================================================================================================
   ' OBJECT IMPLEMENTATION ON TOP OF THE DICT AND EXTENSIONS
   '============================================================================================================

   '-------------------------------------------------------------------------------
   ' Method: MakeClass
   '
   '  Used to build a new bag_object class from an existing one.
   '
   ' Parameters:
   '   - nameToUse - the name for the new class. In practice, usually the same name
   '                 that's given to the global containing object. 
   '   - args      - DictMake compatible arguments which will be copied to the new object
   '   - metaArgs  - DictMake compatible arguments which will be copied to the new object's metadata
   '
   ' Returns:
   '   - a Bag Class object which inherits from the called object
   '
   ' Exceptions:
   '   - Fatal if the called object is not a bag class
   ' 
   ' Usage: 
   '  Dim classNew
   '  Set classNew = classBase.MakeClass "classNew", "self.is_virtual=<eval>False"
   '
   '-------------------------------------------------------------------------------
   Function MakeClass (nameToUse, args, metaArgs)
   
      ' verify the right kinds of thingies
      If NOT objectMetadata("self.is_bag_object") = True Then
         Err.Raise 1, "", "TBD: FATAL: specified objecct to inherit from is not a bag_object class"
      End If

      If objectMetadata("self.is_instance") = True Then
         Err.Raise 1, "", "TBD: FATAL: specified objecct to inherit from is not a bag_object class"
      End If

      ' set up
      Dim newClassToCreate
      'Set args = me.MakeBag(args, metaArgs)
      
      ' make the new one
      Set newClassToCreate = me.MakeBag(NULL, NULL)
      newClassToCreate.ApplyKeys localData
      newClassToCreate.ApplyMetadata objectMetadata
      newClassToCreate.objectMetadata("self.is_bag_object") = True
      
      ' put the new name in (overwrites the old one)
      newClassToCreate.objectMetaData("self.class_name") = nameToUse
      
      ' and inheritence
      Set newClassToCreate.objectMetaData("self.inherits_from") = me
      Set newClassToCreate.objectMetaData("self.inheritence_list") = MakeBag (NULL, NULL) ' objectMetadata("self.inheritence_list")

      ' we can't just copy the whole object, because then we'd just have aa pointer to the same object
      ' so instead we have to do a deeper copy, key by key.
      If objectMetaData.Exists("self.inheritence_list") Then
         If NOT IsEmpty(objectMetaData("self.inheritence_list")) Then
            If objectMetaData("self.inheritence_list").Count > 0 Then
               objectMetaData("self.inheritence_list").CopyTo newClassToCreate.objectMetaData("self.inheritence_list"), NULL
            End If
         End If
      End If

      newClassToCreate.objectMetaData("self.inheritence_list")(me) = True
      newClassToCreate.objectMetaData("self.inheritence_list")(objectMetadata("self.class_name")) = True

      ' and copy anything else into the new object
      newClassToCreate.ApplyKeys args 
      newClassToCreate.ApplyMetaData metaArgs 
   
      Set MakeClass = newClassToCreate
   
   End Function

   '-------------------------------------------------------------------------------
   ' Method: NewObject
   '
   '  Makes an instance of a bag class
   '
   ' Parameters:
   '   - args      - DictMake compatible input, not required (can be NULL) items passed
   '                 in are simply added to the object
   '   - metaArgs  - DictMake compatible arguments which will be copied to the new object's metadata
   '
   ' Returns:
   '   - a new bag dict instance of the class
   '
   ' Exceptions:
   '   - Fatal if this isn't a bag class object
   ' 
   ' Usage:
   '  Dim myInstance
   '  Set myInstance = myClass.NewObject, NULL
   '
   '-------------------------------------------------------------------------------
   Function NewObject (args, metaArgs)
   
      ' verify the right kinds of thingies
      If NOT objectMetadata("self.is_bag_object") = True Then
         Err.Raise 1, "", "TBD: FATAL: specified objecct to inherit from is not a bag_object class"
      End If
      If objectMetadata("self.is_instance") = True Then
         Err.Raise 1, "", "TBD: FATAL: specified objecct to inherit from is not a bag_object class"
      End If
      If objectMetadata("self.is_virtual") Then
         Err.Raise 1, "", "FATAL: CANNOT INSTANTIATE VIRTUAL CLASS"
      End If
      
      Dim theNewObject
      Set theNewObject = me.MakeBag(NULL, NULL)

      theNewObject.ApplyKeys me.localData
      theNewObject.ApplyMetadata me.objectMetadata
      
      'Set args = me.MakeBag(args, NULL)
      theNewObject.ApplyKeys args 
      
      theNewObject.objectMetadata("self.is_instance") = True
      
      If theNewObject.objectMetadata.Exists("method.constructor") Then
         theNewObject.Method "Constructor", args
      End If

      Set NewObject = theNewObject
   
   End Function

   '-------------------------------------------------------------------------------
   ' Method: ApplyKeys
   '
   '  Applies the keys passed to the contained dictionary, without regard to whether
   '  those keys already exist
   '
   ' Parameters (required):
   '   - args - DictMake compatible arguments to be applied
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  newThing.ApplyKeys(oldThing)
   '
   '-------------------------------------------------------------------------------
   Sub ApplyKeys (args)
      CopyKeys args, localData, null
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: ApplyMetaData
   '
   '  Applies the keys passed to the objectMetadata dictionary, without regard to whether
   '  those keys already exist
   '
   ' Parameters (required):
   '   - args - DictMake compatible arguments to be applied
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  newThing.ApplyKeys(oldThing)
   '
   '-------------------------------------------------------------------------------
   Sub ApplyMetaData (args)
      CopyKeys args, objectMetadata, null
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: IsOfClass
   '
   '  Returns True if this is a bag object of the specified class
   '
   ' Parameters:
   '   - classToCheck - can be class object or string name
   '
   ' Returns:
   '   - True/False is me inherits from the specified class at any point in it's 
   '                ancestry
   '
   ' Usage:
   '  myVar = myDict.IsOfClass(classBase)
   '
   ' Notes:
   '
   '-------------------------------------------------------------------------------
   Function IsOfClass(classToCheck)
      IsOfClass = False
         ' verify the right kinds of thingies
      If objectMetadata("self.is_bag_object") = True Then
         IsOfClass = (objectMetadata("self.inheritence_list").Item(classToCheck) = True) ' comparison results in a t/f result
      End If
   End Function


   '-------------------------------------------------------------------------------
   ' Method: ApplyMethod
   '
   '  Used to apply a method to the object. Method keys always start with "method."
   '  and the value is the string name of the subroutine to call
   '
   ' Parameters (required):
   '   - nameToUse - the name of the method (will be made lower case before 
   '                 applying to the key)
   '
   ' Parameters (allowed):
   '   - args key:vector        - string: full name of the subroutine to call. if not 
   '                              provided, one is created (see below)
   '   - args key:required_keys - string of one or more required keys, comma seperated if
   '                              more than one. IMPORTANT: SEE NOTES BEFORE USING!!!
   '
   ' any other specified keys will be applied to the OBJECT METADATA after the keys for 
   ' this property are created
   '
   ' Returns:
   '   - No return value
   '
   ' Exceptions:
   '   - Fatal if not a bag object
   ' 
   ' Usage:
   '  MyClass.ApplyMethod "MethodName", NULL
   '  MyClass.ApplyMethod "MethodName", "vector=>FunctionNameIWantToCall"
   '  (See also notes section below)
   '
   ' Notes:
   '  Method keys in the form "method.<nameToUse>" and default vectors (if none provided)
   '  are in the form <metadata key:class_name>_<nameToUse>
   '
   '  IMPORTANT NOTE ON required_keys READ BEFORE USING: if you send in the key
   '  required_keys, it will create a "custom signature vector". When the Method() entry point
   '  is called, the provided keys will be evaluated against the required_keys list
   '  and if there's a match, the matching vector will be called. This creates a real
   '  form of polymorphism. So you can do this:
   '
   '     myClass.ApplyMethod "Create", "required_keys=>company_name|vector=>MyClass_Create_Company"
   '     myClass.ApplyMethod "Create", "required_keys=>last_name|vector=>MyClass_Create_Person"
   '
   '  And when you do a method call, the provided keys will be evaluated against the required
   '  keys lists and the appropriate vector will be called. If no match is made, and a 
   '  default vector exists, then the default vector will be called. One facility this creates
   '  is the ability to use the default vector to handle errors.
   '
   '  IT IS RECCOMENDED THAT YOU ALWAYS EXPLICITLY SPECIFY THE VECTOR FOR METHODS
   '  VECTORED TO BASED ON CUSTOM KEYS. If you don't do this, a new vector will be
   '  automatically generated, based on the number of custom vectors already existing.
   '  The custom vector will be in the form:
   '    ClassName_method_<nameToUse>_Custom_<number of custom signatures already existing for this nameToUse+1>
   '
   '-------------------------------------------------------------------------------
   Sub ApplyMethod (nameToUse, args)
      VectorSet "method", "", nameToUse, args 
   End Sub
   
   '-------------------------------------------------------------------------------
   ' Method: Method
   '
   '  calls a method (a function associated with a bag_class object)
   '
   ' Parameters (required):
   '   - nameToUse   - string name of the method to call
   '   - args        - DictMake compatible input (can be NULL)
   '
   ' Returns:
   '   - whatever return value the called method returns
   '
   ' Usage:
   '   not caring about return value: myObject.Method "MethodToCall", NULL
   '   capturing non object return:   foo = myObject.Method("MethodToCall", "key_to_pass_to_method_code=>foo")
   '   capturing object return value: Set foo = myObject.Method("MethodToCall", Array("key_to_pass_to_method_code",foo))
   '
   '-------------------------------------------------------------------------------
   Function Method(nameToUse, args)
      Method = NULL
      BagDictCreate args, NULL

      If NOT (objectMetadata("self.is_instance") = True) Then
         Err.Raise 1, "", "TBD: FATAL: CANNOT CALL INTO AN ITEM WHICH IS NOT AN INSTANCE"
      End If
      
      Dim vector, stringCommand, evalResult
      vector = VectorGet("method", "", nameToUse, args)
      
      If vector = "" Then
         Err.Raise 1, "", "TBD: FATAL: METHOD " & nameToUse & " IS NOT DEFINED"
      End If
      
      stringCommand = vector & " (me, args)"
      evalResult = VectorCall(stringCommand, args) ' returns an array of 1 element
      
      If NOT (IsNull(evalResult) OR IsEmpty(evalResult)) Then
         If IsObject(evalResult(0)) Then
            Set Method = evalResult(0)
         Else
            Method = evalResult(0)
         End If
      End If

   End Function
   
   '-------------------------------------------------------------------------------
   ' Method: ApplyProp
   '
   '  Used to apply a property to the object. Property keys always start with "prop_<get/set/let>_"
   '  and the value is the string name of the function to call
   '
   ' Parameters (required):
   '   - nameToUse - the name of the property (will be made lower case before 
   '                 applying to the key)
   '   - direction - the string "set", "let", or "get" - the type of function to call
   '   - args      - anything else, DictMake compatible input (can be NULL)
   '
   ' Parameters (allowed in args):
   '   - args key:vector        - string: full name of the function to call. if not 
   '                              provided, one is created (see below)
   '   - args key:return_type   - string: indicates the type of data this call returns
   '   - args key:required_keys - string of one or more required keys, comma seperated if
   '                              more than one. IMPORTANT: SEE NOTES BEFORE USING!!!
   '
   ' any other specified keys will be applied to the OBJECT METADATA after the keys for 
   ' this property are created
   '
   ' Returns:
   '   - No return value
   '
   ' Exceptions:
   '   - Fatal if not a bag object
   ' 
   ' Usage:
   '  MyClass.ApplyProp "PropertyName", "get", NULL
   '  MyClass.ApplyProp "PropertyName", "get", "return_type=>string"
   '  MyClass.ApplyProp "PropertyName", "get", "vector=>FunctionNameIWantToCall"
   '  MyClass.ApplyProp "PropertyName", "get", "vector=>FunctionNameIWantToCall|return_type=>String"
   '
   ' Notes:
   '  Prop keys are stored in the form "prop.<nametouse>_<get/set>" and default vectors (if none provided)
   '  are in the form <object key:self.class_name>_<nametouse>_<get/set> (let is converted to set)
   '  lets and sets are done in a result type sensitive way, so seperate code is not needed for both
   '
   '  If you specify key:return_type, it is replaced with prop_<nametouse>_<get/set>.return_type
   '
   '  IMPORTANT NOTE ON required_keys READ BEFORE USING: if you send in the key
   '  required_keys, it will create a "custom signature vector". When the Prop() entry point
   '  is called, the provided keys will be evaluated against the required_keys list
   '  and if there's a match, the matching custom signature vector will be called. This creates a real
   '  form of polymorphism. So you can do this:
   '
   '     myClass.ApplyProp "Create", "required_keys=>company_name|vector=>MyClass_Create_Company"
   '     myClass.ApplyProp "Create", "required_keys=>last_name|vector=>MyClass_Create_Person"
   '
   '  And when you do a prop call, the provided keys will be evaluated against the required
   '  keys lists and the appropriate vector will be called. If no match is made, and a 
   '  default vector exists, then the default vector will be called. One facility this creates
   '  is the ability to use the default vector to handle errors.
   '
   '  IT IS RECCOMENDED THAT YOU ALWAYS EXPLICITLY SPECIFY THE VECTOR FOR METHODS
   '  VECTORED TO BASED ON CUSTOM SIGNATURES. If you don't do this, a new vector will be
   '  automatically generated, based on the number of custom vectors already existing.
   '  The custom vector will be in the form:
   '    ClassName_prop_<nameToUse>_<direction>_Custom_<number of custom signatures already existing for this nameToUse & direction+1>
   '
   '-------------------------------------------------------------------------------
   Sub ApplyProp (nameToUse, direction, args)
      if lCase(direction) = "let" Then
         direction = "set"
      End If
      VectorSet "prop", direction, nameToUse, args
   End Sub

   '-------------------------------------------------------------------------------
   ' Property: Prop (Get/Set/Let)
   '
   '  calls a prop (a property type function set associated with a bag_class object)
   '  (indirectly, via CallProp)
   '
   ' Parameters (required):
   '   - nameToUse   - string name of the property to call
   '   - args        - DictMake compatible input
   '
   ' Parameters (special):
   '   - value       - for setter properties only
   '
   ' Returns:
   '   - whatever return value the called method returns
   '
   ' Usage:
   '   (start example)
   '   foo = myObject.Prop("PropToCall", "key_to_pass_to_method_code=>foo")
   '   myObject.Prop("PropToCall", NULL) = "foo"
   '   foo = (myObject.Prop("PropToCall", NULL) = "foo") 
   '   (end)
   '
   '-------------------------------------------------------------------------------
   Property Get Prop(nameToUse, args)
      Dim result
      result = CallProp(nameToUse, "get", args, NULL)
      If NOT (IsNull(result) OR IsEmpty(result)) Then
         If IsObject(result(0)) Then
            Set Prop = result(0)
         Else
            Prop = result(0)
         End If
      End If
   End Property
   
   Property Set Prop(nameToUse, args, value)
      Dim result
      result = CallProp(nameToUse, "set", args, NULL)
      If NOT (IsNull(result) OR IsEmpty(result)) Then
         If IsObject(result(0)) Then
            Set Prop = result(0)
         Else
            Prop = result(0)
         End If
      End If
   End Property
   
   Property Let Prop(nameToUse, args, value)
      Dim result
      result = CallProp(nameToUse, "set", args, NULL)
      If NOT (IsNull(result) OR IsEmpty(result)) Then
         If IsObject(result(0)) Then
            Set Prop = result(0)
         Else
            Prop = result(0)
         End If
      End If
   End Property

   '============================================================================================================
   ' PRIVATE INTERNAL METHODS USED BY THE ApplyMethod, ApplyProp, Method and Prop CALLS
   '============================================================================================================

   '-------------------------------------------------------------------------------
   ' Method: CallProp (Private)
   '
   '  calls a prop (a property type function set associated with a bag_class object)
   '
   ' Parameters (required):
   '   - nameToUse   - string name of the property to call
   '   - direction   - string of "get", "set" or "let"
   '   - args        - DictMake compatible input
   '   - value       - null for getter methods
   '
   ' Returns:
   '   - an array of one element which is whatever return value the called method returns
   '
   ' Usage:
   '   Property Let Prop(nameToUse, args, value)
   '     Assign Prop,CallProp(nameToUse, "let", args, value)
   '  End Property

   '-------------------------------------------------------------------------------   
   Private Function CallProp (nameToUse, direction, args, value)

      CallProp = NULL
      BagDictCreate args, NULL

      If NOT (objectMetadata("self.is_instance") = True) Then
         Err.Raise 1, "", "TBD: FATAL: CANNOT CALL INTO AN ITEM WHICH IS NOT AN INSTANCE"
      End If
      
      Dim vector, stringCommand
      vector = VectorGet("prop", direction, nameToUse, args)
      
      If vector = "" Then
         Err.Raise 1, "", "TBD: FATAL: METHOD " & nameToUse & " IS NOT DEFINED"
      End If

      if direction = "get" Then
         stringCommand = vector & " (me, args)"
      else
         stringCommand = vector & " (me, args, value)"
      End if

      CallProp = VectorCall(stringCommand, args) ' returns an array of 1 element
      
   End Function 

   '-------------------------------------------------------------------------------
   ' Method: VectorCalculationCommon (Private)
   '
   '  Used internally to set up data for the other vector calculation code
   '
   ' Parameters (required):
   '   - fullName             - this one gets modified internally 
   '   - methodOrProp         - this one gets made lower case
   '   - directionForProperty - this one gets validated and made lower case
   '   - nameToUse            - this one gets made lower case
   '   - args                 - this one gets made into a bag dict
   '
   '-------------------------------------------------------------------------------   
   Private Sub VectorCalculationCommon(fullName, methodOrProp, directionForProperty, nameToUse, args)
      ' verify the right kinds of thingies
      If NOT objectMetadata("self.is_bag_object") = True Then
         Err.Raise 1, "", "TBD: FATAL: specified objecct is not a bag_object"
      End If

      ' set up the arguments
      nameToUse = lcase(nameToUse)
      methodOrProp = lcase(methodOrProp)
      directionForProperty = lcase(directionForProperty)
      Set args = MakeBag(args, NULL)
   
      ' verify "method" or "prop"
      If NOT (methodOrProp = "method" OR methodOrProp = "prop") Then
         Err.Raise 1, "", "TBD: FATAL: Internal error, VectorCalculationCommon passed methodOrProp = " & methodOrProp
      End If
      
      ' look for optional special sub item called "args", move to base item
      If args.Exists("args") Then
         args.ApplyKeys args("args")
         args.Remove "args"
      End If  

      ' now lets get the names sorted out
      fullName = methodOrProp & "." & nameToUse
      If methodOrProp = "prop" Then
         fullName = nameToUse & "_" & directionForProperty
      End If
      
   End Sub
   
   '-------------------------------------------------------------------------------
   ' Method: VectorSet (Private)
   '
   '  Sets a vector based on inputs
   '
   ' Parameters (required):
   '   - methodOrProp         - the string "method" or "prop"
   '   - directionForProperty - if called for a prop, this should be either get or set
   '   - nameToUse            - the name of the method
   '   - args                 - this one gets made into a bag dict
   '
   ' Parameters (allowed):
   '   - key:vector           - the vector to use (overrides automatic generation)
   '   - key:required_keys    - comma seperated list of required keys for custom signature 
   '
   ' Returns:
   '  No return value
   '
   ' Notes:
   '  Manages custom signatures, automatic vector generation, etc.
   '
   '-------------------------------------------------------------------------------
   Private Function VectorSet(methodOrProp, directionForProperty, nameToUse, args) 

      Dim fullName, propString
      VectorCalculationCommon fullName, methodOrProp, directionForProperty, nameToUse, args 
      propString = ""
      If methodOrProp = "prop"  Then
         propString = directionForProperty & "_"
      End If

      ' were we passed a custom signature?
      If args.Exists("required_keys") Then  ' required_keys is a comma seperated list
         ' if we haven't done this yet, initialize the signature counter
         If NOT objectMetadata.Exists(fullName & ".number_of_custom_signatues") Then
            objectMetadata(fullName & ".number_of_custom_signatues") = 1
         Else
            objectMetadata(fullName & ".number_of_custom_signatues") = objectMetadata(fullName & ".number_of_custom_signatues") + 1
         End If
         
         objectMetadata(fullName & ".custom_signatue_number_" & objectMetadata(fullName & ".number_of_custom_signatues")) = Split(args("required_keys"),",")
         fullName = fullName & ".custom_vector_number_" & objectMetadata(fullName & ".number_of_custom_signatues")
         
         ' now, if they didn't provide us with a vector, we have to generate one
         If NOT args.Exists("vector") Then
            ' IMPORTANT: THIS NEEDS TO BE CAREFULLY DOCUMENTED!
            ' RECCOMEND CUSTOM SIGNATURES ALWAYS HAVE EXPLICIT VECTORS!
            args("vector") = objectMetadata("self.class_name") & "_" & nameToUse & propString & "_Custom_" & objectMetadata(fullName & ".number_of_custom_signatues")
         End If
         
         args.Remove "required_keys"
         
      End If
      
      ' were we passed an explicit vector? if not, build the default one
      If NOT args.Exists("vector") Then
         args("vector") = objectMetadata("self.class_name") & "_" & nameToUse & propString
      End If

      args(fullName) = args("vector")
      args.Remove "vector"
      
      If methodOrProp = "prop" AND args.Exists("return_type") Then
         args(fullName & ".return_type") = args("return_type")
         args.Remove("return_type")
      End If
      
      me.ApplyMetadata args.localData
   
   End Function
   
   '-------------------------------------------------------------------------------
   ' Method: VectorGet (Private)
   '
   '  Gets a vector based on inputs
   '
   ' Parameters (required):
   '   - methodOrProp         - the string "method" or "prop"
   '   - directionForProperty - if called for a prop, this should be either get or set
   '   - nameToUse            - the name of the method
   '   - args                 - this one gets made into a bag dict
   '
   ' Returns:
   '  The vector found or an empty string.
   '
   ' Notes:
   '  Manages custom signatures BASED ON THE CONTENTS OF args
   '  including automatic vector generation, etc.
   '
   '-------------------------------------------------------------------------------
   Private Function VectorGet(methodOrProp, directionForProperty, nameToUse, args) ' returns empty string if no match found, else vector

      ' set up vars and data
      Dim fullName, stringCommand
      VectorCalculationCommon fullName, methodOrProp, directionForProperty, nameToUse, args 
      
      stringCommand = ""

      ' do we have a default vector?
      If objectMetadata.Exists(fullName) Then
         stringCommand = objectMetadata(fullName)
      End If
      
      ' do we possibly have a custom signature vector?
      If objectMetadata.Exists(fullName & ".number_of_custom_signatues") Then

         ' we'll need these
         Dim signatureIndex, currentWordSetArray, currentWordIndex, customSignatureFound
               
         ' first loop thru the set of signatures
         For signatureIndex = 1 to objectMetadata(fullName & ".number_of_custom_signatues")

            ' get the array of words for this signature
            currentWordSetArray = objectMetadata(fullName & ".custom_signatue_number_" & signatureIndex)

            ' let's pretend for a moment that this is the right one, until we're proven wrong
            customSignatureFound = signatureIndex

            ' can we can prove it's not?
            For currentWordIndex = 0 to UBound(currentWordSetArray)
               If NOT args.Exists(currentWordSetArray(currentWordIndex)) Then ' if we can't find the word
                  customSignatureFound = 0                                    ' then we say Not Found
                  currentWordIndex = UBound(currentWordSetArray)              ' and we move on to the next signature
               End If
            Next

            ' now let's see if we still think we have a match
            If NOT customSignatureFound = 0 Then
               ' if we got here, we have a winner!
               fullName = "method." & nameToUse & ".custom_vector_number_" & customSignatureFound
               stringCommand = objectMetadata(fullName) ' & " (me, args)"
            End If
         Next
      End If
      
      VectorGet = stringCommand
   End Function

   '-------------------------------------------------------------------------------
   ' Method: VectorCall (Private)
   '
   '  calls a vector based on inputs
   '
   ' Parameters (required):
   '   - stringCommand        - the vector to call
   '   - args                 - the args to be passed in to the vector
   '
   ' Returns:
   '  No return value
   '
   ' Notes:
   '  Manages custom signatures, automatic vector generation, etc.
   '
   '-------------------------------------------------------------------------------
   Private Function VectorCall(byVal stringCommand, args)

      ' args has to be here because it gets called in the Eval

      Err.Clear
      On Error Resume Next
      ' JumpStation
      ' set breakpoint below here
      ' set watch on the variable stringCommand
      ' set target routine breakpoint based on the contents of stringCommand
      VectorCall = Array(Eval(stringCommand))
            
      If Err.Number > 0 Then
         PRINT "TBD: ERROR (" & err.number & " " & err.description & " in " & Err.Source & ") ENCOUNTERED CALLING " & stringCommand
         'PRINT Err.Helpfile & " topic for the following HelpContext: " & Err.HelpContext
         'Err.Clear
      End If

      'Err.Clear
      On Error Goto 0

   End Function
   
   

   '============================================================================================================
   ' Now we need add all functions that the Dictionary already supports
   ' For documentation, see docunemtation for the VBScript dictionary object
   '============================================================================================================

   '-------------------------------------------------------------------------------
   ' Property: HashVal
   '
   '  returns the hash value of the object data
   '
   ' Parameters:
   '   - Text
   '
   ' Returns:
   '   - UNKNOWN
   '
   ' Usage:
   '  f = myDict.HashVal(someText)
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)
   '
   '  I have no idea what it does
   '
   '-------------------------------------------------------------------------------
   Public Property Get HashVal(Text)
      HashVal = localData.HashVal(Text)
   End Property

   '-------------------------------------------------------------------------------
   ' Method: Add
   '
   '  Adds a Key/Value pair to the object data
   '
   ' Parameters:
   '   - KeyToUse - any leagal type, usually string. the key for locating the data
   '   - Item     - the item to use
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  myDict.Add "foo",myDataItem
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)   
   '
   '-------------------------------------------------------------------------------   
   Public Sub Add(ByVal keyToUse, ByVal Item)
      If localData.Exists (keyToUse) Then
         localData.Remove keyToUse
      End If
      localData.Add keyToUse, Item
   End Sub

   '-------------------------------------------------------------------------------
   ' Property: Keys
   '
   '  Returns the array of all the keys FOR THE OBJECT DATA (not the metadata)
   '
   ' Parameters:
   '   - none
   '
   ' Returns:
   '   - An array of all the keys FOR THE OBJECT DATA (not the metadata)
   '
   ' Usage:
   '  keyArray = myDict.Keys
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)   
   '
   '-------------------------------------------------------------------------------   
   Public Function Keys()
      Keys = localData.Keys
   End Function

   '-------------------------------------------------------------------------------
   ' Method: Key
   '
   '  This allows you to rename a key FOR THE OBJECT DATA (not the metadata)
   '
   ' Parameters:
   '   - oldKey - the key name to change from
   '   - newKey - the Key name to change to
   '
   ' Returns:
   '   - no return value
   '
   ' Usage:
   '  object.Key(key) = newkey
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)   
   '
   '-------------------------------------------------------------------------------   
   Public Property Let Key(oldKey, newKey)
      localData.Key(oldKey) = newKey
   End Property

   '-------------------------------------------------------------------------------
   ' Property: Items
   '
   '  Returns array of items FROM THE OBJECT DATA (not the metadata)
   '
   ' Parameters:
   '   - None
   '
   ' Returns:
   '   - array of items FROM THE OBJECT DATA (not the metadata)
   '
   ' Usage:
   '  myArray = myObject.Items
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '-------------------------------------------------------------------------------   
   Public Function Items()
      Items = localData.Items
   End Function

   '-------------------------------------------------------------------------------
   ' Method: Exist/Exists 
   '
   '  Returns true if a key exists IN THE OBJECT DATA (not the metadata)
   '  (both Exist and Exists are implemented so I don't have to remember 
   '  which one this supports)
   '
   ' Parameters:
   '   - keyToUse - The key to check for the existance of IN THE OBJECT DATA (not the metadata)
   '
   ' Returns:
   '   - True if the specified key exists in the object data, else false
   '
   ' Usage:
   '  myAnswer = myObject.Exist("myKey")
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '-------------------------------------------------------------------------------   
   Public Function Exists(keyToUse)
      Exists = localData.Exists(keyToUse)
   End Function
   
   Public Function Exist(keyToUse)
      Exist = localData.Exists(keyToUse)
   End Function

   '-------------------------------------------------------------------------------
   ' Method: RemoveAll
   '
   '  Remove All keys/values IN THE OBJECT DATA (not the metadata)
   '
   ' Parameters:
   '   - none
   '
   ' Returns:
   '   - no return value
   '
   ' Usage:
   '  myObject.RemoveAll
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '-------------------------------------------------------------------------------   
   Public Sub RemoveAll()
      localData.RemoveAll
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: Remove
   '
   '  Remove a specified key FROM THE OBJECT DATA (not the metadata)
   '
   ' Parameters:
   '   - keyToUse - The key to remove
   '
   ' Returns:
   '   - No return value
   '
   ' Usage:
   '  myObject.Remove(keyToUse)
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '-------------------------------------------------------------------------------   
   Public Sub Remove (keyToUse)
      If localData.Exists(keyToUse) Then
         localData.Remove (keyToUse)
      End If
   End Sub

   '-------------------------------------------------------------------------------
   ' Property: Count
   '
   '  Get count of items in THE OBJECT DATA (not the metadata)
   '
   ' Parameters:
   '   - None
   '
   ' Returns:
   '   - Count of items in the object data
   '
   ' Usage:
   '  foo = myObject.Count
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '-------------------------------------------------------------------------------   
   'Get count of items in dictionary
   Public Property Get Count()
      Count = localData.Count
   End Property

   '-------------------------------------------------------------------------------
   ' Property: CompareMode
   '
   '  Used to set and get the CompareMode flag for the object data
   '
   ' Parameters:
   '   - for the set method, the new compare mode
   '
   ' Returns:
   '   - from the Get method, the current compare mode
   '
   ' Usage:
   '  
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '  From the MS docs: Acceptable values are 0 (Binary), 1 (Text), 2 (Database). 
   '  Values greater than 2 can be used to refer to comparisons using specific 
   '  Locale IDs (LCID). 
   '
   '-------------------------------------------------------------------------------   
   'Get Property for CompareMode
   Public Property Get CompareMode()
      CompareMode = localData.CompareMode
   End Property

   'Let Property for CompareMode
   Public Property Let CompareMode(newMode)
      localData.CompareMode = newMode
   End Property

   '-------------------------------------------------------------------------------
   ' Property: Item
   '
   '  Item is the Default property for dictionary. This call gets/sets items in the
   '  object data only, not the metadata. 
   '
   ' Parameters:
   '   - keyToUse - (Get and Set) the key to use
   '   - Value    - (Set only) the value to set for the key
   '
   ' Returns:
   '   - Get returns the value associated with the key
   '
   ' Usage:
   '  Set: myDict(keyToUse) = Value
   '  Get: foo = myDict(keyToUse)
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '  Item is the Default property for dictionary. So we need to use default keyword with 
   '  Property Get... Default keyword can be used with a only one Function or Get Property
   '
   '-------------------------------------------------------------------------------   
   Public Default Property Get Item(keyToUse)
      'If a object is stored for the Key
      'then we need to use Set to return the object
      Item = Empty
      If NOT IsObject(keyToUse) Then
         If keyToUse = "self.is_bag_dict" Then
            Item = True
         End If
      End If
      If IsEmpty(Item) Then
         If localData.Exists (keyToUse) Then
            If IsObject(localData.Item(keyToUse)) Then
               Set Item = localData.Item(keyToUse)
            Else
               Item = localData.Item(keyToUse)
            End If
         End If
      End If
   End Property

   Public Property Let Item(keyToUse, Value)
      'Check of the value is an object
      If IsObject(Value) Then
         'The value is an object, use the Set method
         Set localData(keyToUse) = Value
      Else
         'The value is not an object assign it
         localData(keyToUse) = Value
      End If
   End Property

   Public Property Set Item(keyToUse, Value)
      If IsObject(Value) Then
         'The value is an object, use the Set method
         Set localData(keyToUse) = Value
      Else
         'The value is not an object assign it
         localData(keyToUse) = Value
      End If
   End Property

   '-------------------------------------------------------------------------------
   '  Extensions by Tarun Lalwani
   '  (http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)
   '-------------------------------------------------------------------------------
   
   '============================================================================================================
   ' Method: AddFromDictionary
   '
   '  Copies the keys/values from another dict to this one
   '
   ' Parameters:
   '   - oldDict - the dictionary object (or bag dict) to copy from
   '
   ' Returns:
   '   - no return value
   '
   ' Usage:
   '  myDict.AddFromDictionary(theOtherDict)
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '  Copies all keys. 
   '
   '-------------------------------------------------------------------------------
   Public Sub AddFromDictionary(oldDict)
      aKeys = oldDict.Keys

      For Each sKey In aKeys
         If IsObject(oldDict(sKey)) Then
            Set localData(sKey) = oldDict(sKey)
         Else
            localData(sKey) = oldDict(sKey)
         End If
      Next
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: LoadFromDictionary
   '
   '  REPLACES the *OBJECT DATA* in this dictionary with that of another dictionary
   '
   ' Parameters:
   '   - oldDict - the dictionary to copy from
   '
   ' Returns:
   '   - no return value
   '
   ' Usage:
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '  Similar to AddFromDictionary, this one removes existing data first
   '
   '-------------------------------------------------------------------------------
   Public Sub LoadFromDictionary(oldDict)
      localData.RemoveAll
      Me.AddFromDictionary oldDict
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: FileReadDict
   '
   '  Attempts to populate this from a disk image - ONLY POPULATES OBJECT DATA
   '  AND ONLY WITH STRING DATA. Adds data from disk image to existing data in the object
   '
   ' Parameters:
   '   - FileName  - the name of the file
   '   - Delimiter - the key/value seperator
   '
   ' Returns:
   '   - no return value
   '
   ' Usage:
   '  myObject.FileReadDict myFileName, "=>"
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '-------------------------------------------------------------------------------
   Public Sub FileReadDict(FileName, Delimiter)
      Set FSO = CreateObject("Scripting.FileSystemObject")
      Set oFile = Fso.OpenTextFile (FileName)

      'Read the file line by line
       While Not oFile.AtEndOfStream
         sLine = oFile.ReadLine
         KeyValue = Split(sLine, Delimiter)
         localData(KeyValue(0)) = KeyValue(1)
      Wend

      Set oFile = Nothing
      Set FSO = Nothing
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: LoadFromFile
   '
   '  Attempts to populate this from a disk image - ONLY POPULATES OBJECT DATA
   '  AND ONLY WITH STRING DATA. *REPLACES* data in existing object with data from disk
   '
   ' Parameters:
   '   - FileName  - the name of the file
   '   - Delimiter - the key/value seperator
   '
   ' Returns:
   '   - no return value
   ' 
   ' Usage:
   '  myObject.LoadFromFile myFileName, "=>"
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '-------------------------------------------------------------------------------
   Public Sub LoadFromFile(FileName, Delimiter)
      localData.RemoveAll
      Me.AddFromFile FileName, Delimiter
   End Sub

   '-------------------------------------------------------------------------------
   ' Method: FileWriteDict
   '
   '  Attempts to write this object's OBJECT DATA to a file. Really can only output
   '  string data. 
   '
   ' Parameters:
   '   - FileName  - the name of the file
   '   - Delimiter - the key/value seperator
   '
   ' Returns:
   '   - no return value
   ' 
   ' Usage:
   '  myObject.FileWriteDict myFileName, "=>"
   '
   ' Notes:
   '  This was part of the dictionary passthrough code was borrowed from Tarun Lalwani
   '  (published at: http://knowledgeinbox.com/articles/vbscript/extending-dictionary-object/)      
   '
   '-------------------------------------------------------------------------------
   Public Sub FileWriteDict(FileName, Delimeter)
      Set FSO = CreateObject("Scripting.FileSystemObject")
      Set oFile = FSO.CreateTextFile(FileName, True)

      Dim aKeys
      aKeys = localData.Keys

      'Write the key value pairs line by line
      For Each sKey In aKeys
         oFile.WriteLine sKey & Delimeter & objectMetadata(sKey)
      Next

      'Close the file
      oFile.Close

      Set oFile = Nothing
      Set FSO = Nothing
   End Sub

End Class

'============================================================================================================
' OBJECT MODEL BASEMOST CLASS
'============================================================================================================

'-------------------------------------------------------------------------------
' Bag Class: classBase
'
'  Base-most virtual class, all other things inherit from this. Implements basic data 
'  required for operation of the bag dictionaries as class/instance objects with
'  inheritance, polymorphism, and so on. 
'
' Inherits: 
'  Nothing
' 
' Implements:
'   - key:self.class_name => classBase
'   - key:self.is_virtual => True
'
'-------------------------------------------------------------------------------
Dim classBase
Set classBase = BagDictMake(NULL, NULL)
classBase.ApplyMetadata "self.class_name=>classBase"
classBase.ApplyMetadata "self.is_virtual=><eval>True"
classBase.ApplyMetadata "self.is_instance=><eval>False"
classBase.ApplyMetadata "self.is_bag_object=><eval>True"
classBase.ApplyMethod "Destroy", NULL

'-------------------------------------------------------------------------------
' Bag Method: Destroy (ClassBase_Destroy)
'
'  Root most destructor
'
' Parameters: 
'   - none
'
' Returns:
'   - no return value
'
' Usage:
'  myBagObject.Method "Destroy", NULL
'
' Notes:
'   Placeholder
'
'-------------------------------------------------------------------------------
Function ClassBase_Destroy (self, args)
   ' placeholder, so that there is a destructor at this level
End Function


'============================================================================================================
' OBJECT MODEL FILE CLASSES
'============================================================================================================

'-------------------------------------------------------------------------------
' Bag Class: classFileSystemItem
' 
'  Virtual class, starting point for all file system objects
'
' Inherits: 
'  classBase
' 
' Implements: 
'  Nothing (yet - AMM 6/15/2008)
'
'-------------------------------------------------------------------------------
Dim classFileSystemItem
Set classFileSystemItem = classBase.MakeClass("classFileSystemItem",NULL, "self.is_virtual=><eval>True") ' self.is_virtual=False isn't needed, as it's inherited, but it's clearer this way
Dim ClassFileSystemItem_FSO
Set ClassFileSystemItem_FSO = CreateObject("Scripting.FileSystemObject")
' this is sort of a defacto constructor, and why we have a destructor
classFileSystemItem.ApplyKeys Array("fso",ClassFileSystemItem_FSO)

'-------------------------------------------------------------------------------
' Bag Method: Destroy  (classFileSystemItem_Destroy)
'
'  Readys the class for being destroyed. This class creates a default "fso" entry
'  on construction of the class. This is part of the object data. This destroy method
'  removes that.
'
' Overrides: 
'  ClassBase_Destroy
'
' Parameters: 
'   - none
'
' Returns:
'   - no return value
'
' Usage:
'  myBagInstance.Method "Destroy", NULL
'
'-------------------------------------------------------------------------------
classFileSystemItem.ApplyMethod "Destroy", NULL
Function ClassFileSystemItem_Destroy (self, args)
   If NOT IsEmpty(self("file_system_object")) Then
      self("file_system_object").Close
   End If
End Function

'-------------------------------------------------------------------------------
' Bag Class: classFile
' 
'  Refers to files (not directories), not virtual
'
' Inherits: 
'  classBase
'
' Implements:
'   - key:file_name  - name and path to file REQUIRED FOR ALL FILE IO
'   - prop:exist - returns boolean of file existance
'   - method:delete  - deletes file if it exists
'  
'-------------------------------------------------------------------------------
Dim classFile
Set classFile = classFileSystemItem.MakeClass("classFile",NULL,"self.is_virtual=><eval>False") 

'-------------------------------------------------------------------------------
' Bag Property: Exist (ClassFile_Exist)
'
'  Returns true or false if the file specified in key:file_name exists
'
' Parameters: 
'   - none
'
' Returns:
'  true or false if the file specified in key:file_name exists
'
' Usage:
'  myResult = myFileInstance.Prop("Exist", NULL)
'
'-------------------------------------------------------------------------------
Function ClassFile_Exist(self,args)
   self.VerifyHasKeys "file_name", "TBD: FATAL: ClassFile_Exist fails becase self does not have key:file_name"
   ClassFile_Exist = ClassFileSystemItem_FSO.FileExists(self("file_name"))
End Function
classFile.ApplyProp "Exist","Get","vector=>ClassFile_Exist|return_type=>Boolean"

'-------------------------------------------------------------------------------
' Bag Method: Delete (ClassFile_Delete)
' 
'  Deletes the file specified in key:file_name if it exists (else do nothing)
'
' Parameters: 
'   - none
'
' Returns:
'   - no return value
'
' Usage:
'  myFileInstance.Method "Delete", NULL
'
'-------------------------------------------------------------------------------
Function ClassFile_Delete(self,args)
   self.VerifyHasKeys "file_name", "TBD: FATAL: ClassFile_Delete fails becase self does not have key:file_name"
   If ClassFile_Exist(self, args) Then
      ClassFileSystemItem_FSO.DeleteFile(self("file_name"))
   End If
End Function
classFile.ApplyMethod "Delete",NULL

'-------------------------------------------------------------------------------
' Bag Property: MD5 (ClassFile_MD5)
'
'  Returns string of the MD5 for the file specified in key:file_name if it exists
'
' Parameters: 
'   - none
'
' Returns:
'  string of the MD5 for the file specified in key:file_name if it exists
'
' Usage:
'  myResult = myFileInstance.Prop("MD5", NULL)
'
'-------------------------------------------------------------------------------
Function ClassFile_MD5(self,args)
   self.VerifyHasKeys "file_name", "TBD: FATAL: ClassFile_MD5 fails becase self does not have key:file_name"
   ClassFile_MD5 = GetMD5ForFile(self("file_name"))
End Function
classFile.ApplyProp "md5","get","vector=>ClassFile_MD5|return_type=>string"

'-------------------------------------------------------------------------------
' Bag Method: UnitTests (ClassFile_UnitTests)
'
'  Unit tests for methods implemented in ClassFile
'
' Parameters:
'   - none
'
' Returns:
'   - no return value
'
' Usage:
'  myFileInstance.Method "UnitTests", NULL
'
' Notes:
'  This is a minimal unit test as of this time. you have to create a file called
'  C:\foo.txt to use this. Additional tests and automatic creation of test files
'  needs to be created.
'-------------------------------------------------------------------------------
Function ClassFile_UnitTests(self,args)
   Dim myFile
   Set myFile = classFile.NewObject ("file_name=>c:\foo.txt",NULL)
   print myFile.Prop ("Exist", NULL)
   myFile.Method "Delete", NULL
   print "File still exist: " & myFile.Prop ("Exist", NULL)
End Function
classFile.ApplyMethod "UnitTests",NULL

'-------------------------------------------------------------------------------
' Bag Class: classTextFile
' 
'  Refers to text files
'
' Inherits: 
'  classFile
'
' Implements:
'   - key:data       - dict contents of the text file in the Items
'   - method:write   - writes file to disk, overwrites any existing file
'   - method:read    - reads file (if it exists) into key:data, destroys any existing contents
'
' Planned:
'   - method:merge   - reads file specified in args_key:file_name and appends to end of key:data
'   - method:append  - writes file to disk, appending to file specified in args_key:file_name
'   - prop:dirty - returns boolean of whether there exist unsaved changes
'
'-------------------------------------------------------------------------------
Dim classTextFile
Set classTextFile = classFile.MakeClass("classTextFile",NULL,"self.is_virtual=><eval>False") 

'-------------------------------------------------------------------------------
' Bag Method: Constructor (ClassTextFile_Constructor)
'
'  Adds the key:data member (a scripting dictionary)
'
' Parameters:
'   - none
'
' Returns:
'   - no return value
'
' Usage:
'  No need, constructors are called automatically
'
'-------------------------------------------------------------------------------
Function ClassTextFile_Constructor(self,args)
   Set self("data") = CreateObject("Scripting.Dictionary")
End Function
classTextFile.ApplyMethod "Constructor",NULL

'-------------------------------------------------------------------------------
' Bag Method: Write (ClassTextFile_Write)
'
'  writes the contents of bagdict self("data").Items
'  to the file specified in localData.key:file_name
'  overwrites any existing file
'
' Parameters:
'   - Reqd ObjData key:data - a scripting dictionary containing the data
'
' Returns:
'   - no return value
'
' Usage:
'  myTextFileObject.Method "Write",NULL
'
'-------------------------------------------------------------------------------
Function ClassTextFile_Write(self,args)

   Dim fileObject, itemList, currentItem
   If self.Exists("os.file_handle") Then
      self("os.file_handle").Close
   End If
   Set fileObject = ClassFileSystemItem_FSO.CreateTextFile(self("file_name"), True)
   
   itemList = self("data").Items

   'Write the key value pairs line by line
   For Each currentItem In itemList
      fileObject.WriteLine currentItem
   Next

   'Close the file
   fileObject.Close
   Set fileObject = Nothing
         
End Function
classTextFile.ApplyMethod "Write",NULL

'-------------------------------------------------------------------------------
' Bag Method: Read (ClassTextFile_Read)
'
'  Reads the contents of the file specified in localData.key:file_name
'  into self("data").Items (keys are row numbers, starting with 1)
'  overwrites any existing internal data
'
' Parameters:
'   - none
'
' Returns:
'   - no return value
'
' Usage:
'  myTextFileObject.Method "read",null 
'
'-------------------------------------------------------------------------------
Function ClassTextFile_Read(self,args)

   Dim fileObject, i

   ' first zorch the old data
   Set self("data") = BagDictMake(NULL,NULL)
   
   ' next, see if we have a valid file
   If self.Exists("os.file_handle") Then
      self("os.file_handle").Close
      self.Remove("os.file_handle")
   End If
   
   If self.Prop("Exist",NULL) Then
      Set fileObject = ClassFileSystemItem_FSO.OpenAsTextStream(self("file_name"), 1, False)
      Set self("os.file_handle") = fileObject
   
      i = 1
      'while not eof
      Do While fileObject.AtEndOfStream <> True
         self("data")(i) = fileObject.ReadLine
         i = i + 1
      Loop

      If self.Exists("os.file_handle") Then
         self("os.file_handle").Close
         self.Remove("os.file_handle")
      End If

   End If

End Function
classTextFile.ApplyMethod "Read",NULL

'-------------------------------------------------------------------------------
' Bag Method: UnitTests (ClassTextFile_UnitTests)
'
'  Unit tests for methods implemented in ClassTextFile (Overrides the ones in
'  classFile)
'
' Parameters:
'   - none
'
' Returns:
'   - no return value
'
' Usage:
'  Intended usage: ClassTextFile_UnitTests NULL,NULL
'
'-------------------------------------------------------------------------------
Function ClassTextFile_UnitTests(self,args)
   Dim myFile
   Set myFile = classTextFile.NewObject ("file_name=>c:\foo.txt",NULL)

   Dim i
   For i = 1 to 10
      myFile("data")(i) = "" & i & ""
   Next
   
   print "Does it exist on entry: " & myFile.Prop ("Exist", NULL)
   myFile.Method "Write",NULL
   print "Does it exist after write: " & myFile.Prop ("Exist", NULL)
   myFile.Method "Delete", NULL
   print "Does it exist after delete: " & myFile.Prop ("Exist", NULL)
End Function
classTextFile.ApplyMethod "UnitTests",NULL

' UNIT TESTS:
'ClassTextFile_UnitTests NULL,NULL
'ExitAction

'-------------------------------------------------------------------------------
' Bag Class: classINIFile
' 
'  Refers to ini files
'
' Inherits: classFile
'
' Implements:
'   - key:data       - dict contents of the text file
'   - prop:get   - reads value from ini file
'   - method:put     - writes value to ini file
'
'-------------------------------------------------------------------------------
Dim classINIFile
Set classINIFile = classTextFile.MakeClass("classTextFile",NULL,"self.is_virtual=><eval>False") 
On Error Resume Next
Extern.Declare micInteger,"GetPrivateProfileStringA", "kernel32.dll","GetPrivateProfileStringA", micString, micString, micString, micString+micByRef, micInteger, micString 
Extern.Declare micInteger ,"WritePrivateProfileString","Kernel32.dll","WritePrivateProfileStringA",micString,micString,micString,micString
On Error Goto 0

'-------------------------------------------------------------------------------
' Bag Method: Put (classINIFile_Put)
'
'  puts a value into an INI file
'
' Parameters: 
'   - key:section
'   - key:key
'   - key:value
'
' Returns:
'   - boolean of whether the action was successful
'
' Usage:
'  myObject.Method "Put","section=>myIniSection|key=>myKey|value=>myValue"
'
'-------------------------------------------------------------------------------
Function ClassINIFile_Put(self, args)
   Dim SetINIValue
	SetINIValue = Extern.WritePrivateProfileString(args("section"),args("key"),args("value"), self("file_name"))
   classINIFile_Put = True
   if SetINIValue = 0 Then
      classINIFile_Put = False
   End If
End Function
classINIFile.ApplyMethod "put", NULL

'-------------------------------------------------------------------------------
' Bag Property: Get (classINIFile_Get)
'
'  Gets a value from the ini file
'
' Parameters: 
'   - key:section
'   - key:key
'
' Returns:
'   - the string value found in the ini file or empty string
'
' Usage:
'  myData = myINIObject.Prop("get","section=>myIniSection|key=>myKey"
'
'-------------------------------------------------------------------------------
Function ClassINIFile_Get(self, args)
	Dim key, i, key2 
	key = String(255, "-") 'makes a buffer
	i = Extern.GetPrivateProfileStringA(args("section"),args("key"),"", key, 255, self("file_name")) 
	key2 = Left(key,i) 
   classINIFile_Get = key2
End Function
classFile.ApplyProp "get","get","vector=>classINIFile_Get|return_type=>string"

'-------------------------------------------------------------------------------
' Bag Class: classQueue
' 
'  queues contain arrays of objects, and are typically FIFO when being "popped" (as opposed to stacks)
'
' Inherits: ClassBase
'
' Implements:
'   - method:push
'   - method:pop
'
'-------------------------------------------------------------------------------
Dim classQueue
Set classQueue = classBase.MakeClass("classQueue",NULL,"self.is_virtual=><eval>False")

'-------------------------------------------------------------------------------
' Bag Method: Push (ClassQueue_Push)
'
'  Pushes an item onto the stack
'
' Parameters:
'   - passed key:item - the item to push
'
' Returns:
'   - no return value
'
' Usage:
'  myQueueInstance.Method "Push", Array("item",myThingToPush)
'
'-------------------------------------------------------------------------------
Function ClassQueue_Push(self, args)
   ClassQueue_Push = NULL
   ' get the data ready and validate it
   DictMakeExpectItem args, NULL
   args.VerifyHasKeys "item", "TBD: FATAL: ClassQueue_Push called without args.key:item"

   ' has the container been initialized yet? if not, do it!
   If IsEmpty(self.Item("data")) Then
      Set self.Item("data") = self.MakeBag(NULL, NULL) ' DATA IS DEPRICATED
      self.Item("last_counter") = 0
   End If

   ' were we given a name for the item?
   Dim nameToUse
   nameToUse = "" & self.Item("last_counter") & ""
   If args.Exists("name") Then
      nameToUse = args("name")
   End If
   
   ' now add it
   self.Item("data").Item(nameToUse) = args("item")
   self.Item("last_counter") = self.Item("last_counter") + 1
   DictUnwrapItemIfExists args
End Function
classQueue.ApplyMethod "Push", NULL 

'-------------------------------------------------------------------------------
' Bag Property: Pop (ClassQueue_Pop)
'
'  Pops an item off of the stack (FIFO)
'
' Parameters:
'   - none
'
' Returns:
'   - whatever was on top of the stack or the keyed string "<NULL>" (you might
'     push a null on to the stack on purpose, you're less likely to push that on)
'
' Usage:
'  Assign poppedFromStack = myQueueInstance.Prop ("Pop", NULL)
' 
'-------------------------------------------------------------------------------
Function ClassQueue_Pop(self, args)
   BagDictCreate args, NULL 
   ClassQueue_Pop = "<NULL>"
   
   If self.Exists("data") Then
      Dim nameToUse
      nameToUse = self.Item("data").KeyIndex(0)
      If args.Exists("name") Then
         nameToUse = args("name")
      End If
   
      If self.Item("data").Count > 0 Then
         Assign ClassQueue_Pop, self.Item("data").Item(nameToUse)
         self.Item("data").Remove nameToUse
      End If
   End If
End Function
classFile.ApplyProp "Pop","get","vector=>ClassQueue_Pop|return_type=>anytype"




'***********************************************************
' aggregating q:\utils\libs\logging.vbs

'###############################################################################
' Library: Logging.vbs
' 
' THIS IS A KEY ARCHITECTURAL COMPONENT, THE ARCHITECT SHOULD BE NOTIFIED OF 
' CHANGES TO THIS FILE. 
'
' Description:
'  This library provides an advanced logging engine for QTP, Log4VB, loosely
'  based on Log4J.  There are eight levels of log messages, ranging from fatal
'  errors to trace information.
'  See http://en.wikipedia.org/wiki/Log4j and http://logging.apache.org/log4j/1.2/index.html
' 
'  Copyright (C) 2008, 2009, 2010 Akien MacIain
'
'  This program is free software: you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation, either version 3 of the License, or
'  (at your option) any later version.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' Usage:
' 
'  For basic logging, you use the 10 logging functions defined below.  args is a
'  serialized dictionary string that contains two keys, routine and message.  For
'  example, to log a warning from a routine named CheckForValidUserID that there are
'  multiple users with the same ID, you could use a call like this
' 
'  (start code)
'  LogWarn "routine=>CheckForValidUserID|message=>Multiple users found with user id " + userID
'  (end)
' 
'  Note that calling log Fata will cause the test to exit.
'
'  (start code)
'  LogFatal(args)
'  LogFail(args) | LogFailure(args)
'  LogWarning(args) | LogWarn(args)
'  LogInfo(args) | LogPrint(args) 
'  LogDebug(args)
'  LogStatus(args)
'  LogTrace(args)
'  (end)
' 
'  Akien MacIain 2008
'
'###############################################################################
'Option Explicit

'-------------------------------------------------------------------------------
' Function: FrameworkDetectLogging
'   Utility function for the Framework Compilation checking utility
'
' Returns:
'  (integer) always returns 1
'-------------------------------------------------------------------------------
Function FrameworkDetectLogging()
	FrameworkDetectLogging = 1
End Function

'============================================================================================================
' LOGGING PUBLIC INTERFACES AND TESTING

Public Const LOGGING_LEVELS_WIDTH   = 10    ' controls the width of the "levels" column in the output
Public Const LOGGING_LEVEL_FATAL    = 0    ' logging levels numbers (see also below the negative numbers)
Public Const LOGGING_LEVEL_FAIL     = 1
Public Const LOGGING_LEVEL_WARN     = 2
Public Const LOGGING_LEVEL_INFO     = 3
Public Const LOGGING_LEVEL_ACTION   = 4
Public Const LOGGING_LEVEL_DEBUG    = 5
Public Const LOGGING_LEVEL_STATUS   = 6
Public Const LOGGING_LEVEL_TRACE    = 7
Public Const LOGGING_LEVEL_SPECIAL  = 8
Public Const LOGGING_LEVEL_0        = "Fatal"    ' the names for the different levels
Public Const LOGGING_LEVEL_1        = "Failure"
Public Const LOGGING_LEVEL_2        = "Warning"
Public Const LOGGING_LEVEL_3        = "Info"
Public Const LOGGING_LEVEL_4        = "Action"
Public Const LOGGING_LEVEL_5        = "Debug"
Public Const LOGGING_LEVEL_6        = "Status"
Public Const LOGGING_LEVEL_7        = "Trace"
Public Const LOGGING_LEVEL_8        = "Special"

Public Const LOGGING_LEVEL_RESULT  = -1   ' test result at end of test
Public Const LOGGING_LEVEL__1      = "Result"

Public Const LOGGING_LEVEL_PICTURE = -2   ' test result at end of test
Public Const LOGGING_LEVEL__2      = "Picture"

Public Const LOG_FileIO_ForReading = 1, LOG_FileIO_ForWriting = 2, LOG_FileIO_ForAppending = 8        ' constants used for file I/O

'-------------------------------------------------------------------------------
' Bag Class: classmessage
' 
'  is a logging message object. generated by logging entry interfaces, such
'  as LogFatal, LogDebug, etc.
'
' Inherits: 
'  ClassBase
'
' Implements:
'   - object key:message       - the text of the message itself
'   - object key:level         - the level of the message (see the constants for LOGGING_LEVEL_x above
'   - object key:routine       - the name of the routine issuing the call
'   - method:dispatch          - dispatches the message object to all the appenders
'
'-------------------------------------------------------------------------------
Dim classmessage
Set classmessage = classBase.MakeClass("classmessage",NULL,NULL)
classmessage.ApplyMetadata "self.is_virtual=><eval>false"
classmessage.ApplyKeys     "routine=>WAS_NOT_DEFINED"
classmessage.ApplyKeys     "message=>WAS_NOT_DEFINED"

'-------------------------------------------------------------------------------
' Bag Method: Dispatch (ClassMessage_Dispatch)
'
'  Dispacthes the message to all the appenders that have registered with 
'  the loggingConfiguration("appenders.queue")
'
' Parameters:
'   - self("level")                           - the message level
'   - loggingConfiguration("appenders.queue") - the ClassQueue of appenders to work thru
'
' Returns:
'   No return value
'
' Usage:
'  theMessage.Method "Dispatch",NULL
'
' Notes:
'  when the call reaches it's target in the appender, the args will contain
'  one argument, message=>the class message object
'
'-------------------------------------------------------------------------------
Function ClassMessage_Dispatch (self, args)

   Dim i, currentAppenderKey, currentAppender, tempLevel, postMessageFlag
   tempLevel = CInt(self("level"))
   
   If loggingConfiguration("debug_messages") = True Then
      If tempLevel > LOGGING_LEVEL_INFO Then
         tempLevel = LOGGING_LEVEL_INFO
      End If
   End If

   For i = loggingConfiguration("appenders.queue")("data").Count-1 to 0 Step -1 ' this gives us the pointer into the array

      currentAppenderKey = loggingConfiguration("appenders.queue")("data").KeyIndex(i)
      Set currentAppender = loggingConfiguration("appenders.queue")("data")(currentAppenderKey)

      postMessageFlag = currentAppender.Method("DoYouWantMe",Array("message",self))
            
      If postMessageFlag Then
         currentAppender.Method "Write",Array("message",self)
      End If
      
   Next
   
End Function
classmessage.ApplyMethod "Dispatch",NULL

'-------------------------------------------------------------------------------
' Bag Class: classAppender
' 
'  Part of the Log4VB Engine (based loosely on Log4J) - Akien MacIain 2008
'
' Inherits: 
'  ClassBase
'
' Implements:
'   - method:write   - writes a message object to the appender target
'   - object key:name       - name of the appender
'   - object key:level      - default value is zero, all non negative messages
'
'-------------------------------------------------------------------------------
Dim classAppender
Set classAppender = classBase.MakeClass("classAppender", NULL, NULL)
classAppender.ApplyMetadata "self.is_virtual=><eval>false"
classAppender.ApplyKeys     "level=>0"
classAppender.ApplyKeys     "level_override_table_enabled=><eval>False"
classAppender.ApplyKeys     Array("level_override_table",BagDictMake(NULL,NULL))
ClassAppender.ApplyKeys     "enabled=><eval>True" ' this is used to disable an appender

'-------------------------------------------------------------------------------
' Bag Method: Write (ClassAppender_Write)
'
'  Entry point for all appenders to post messages, usually a redirect to another 
'  routine that actually does the work. Default vector is to SimplePrint
'
' Parameters:
'   - args("message") - the message object to post
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "Write", Array("message",mymessageObject)
'
' Notes:
'  Because this is intended to be redirected, it is by default redirected to
'  the ClassAppender_SimplePrint in the class, and is overridden in the instance.
'
'-------------------------------------------------------------------------------
classAppender.ApplyMethod "Write", "vector=>ClassAppender_SimplePrint" ' by default vectors to simple print

'-------------------------------------------------------------------------------
' Bag Method: DoYouWantMe (ClassAppender_DoYouWantMe)
'
'  Routine to determine whether the appender wants the message on offer
'
' Parameters:
'   - args("message") - the message object to post
'
' Returns:
'  Boolean of whether the message meets the criteria of the appender
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: x = myAppenderInstance.Method "DoYouWantMe", Array("message",mymessageObject)
'
' Notes:
'  This is intended to cover the most common cases, but to allow
'  specific appenders to override this call.
'
'-------------------------------------------------------------------------------
Function ClassAppender_DoYouWantMe(self, args)

   Dim postMessageFlag, tempLevel
   
   tempLevel = CInt(args("message")("level"))
   postMessageFlag = False

   If self("enabled") = True Then
      ' check for level override table
      If self("level_override_table_enabled") Then
         If self("level_override_table").Exists(tempLevel) Then
            If self("level_override_table")(tempLevel) = True Then
               postMessageFlag = True
            End If
         End If      
      End If

      ' now check the all flag
      If self("level_override_all") = True Then
         postMessageFlag = True
      End If

      ' check for negative number
      If tempLevel < 0 Then
         ' negatives only work if they're exact matches
         If CInt(self("level")) = tempLevel Then
            postMessageFlag = True
         End If
      Else
         ' positives work if the message is less than or equal to the appender
         If CInt(self("level")) >= tempLevel Then
            postMessageFlag = True
         End If
      End If ' end if tempLevel < 0
   End If ' end if self("enabled")
      
   ClassAppender_DoYouWantMe = postMessageFlag
      
End Function
classAppender.ApplyMethod "DoYouWantMe", "vector=>ClassAppender_DoYouWantMe" 

'-------------------------------------------------------------------------------
' Bag Property: Format (ClassAppender_Format) (Get)
'
'  Entry point for all appenders to post messages, usually a redirect to another 
'  routine that actually does the work. 
'
' Parameters (required):
'   - args("message")("routine") - the routine calling
'   - args("message")("message") - the message being posted
'
' Parameters (allowed):
'   - args("format")            - 3rd priority format override specified by the write routine that called this formatter
'   - self("format")            - 2nd priority format override specified to the appender at instantiation
'   - args("message")("format") - 1st priority format override specified by the initial caller
'   - self("indent_enabled")    - NOT YET IMPLEMENTED (as of Aug 26 2009)
'
' Usage:
'  myFormattedMessage = self.Prop ("Format",args)
'
' Notes:
'  Valid formatting strings are usually in the method, and are usually in the form:
'  \yyyy\mm\dd \hh\mn\ss \ts \level \indent \routine \message
'
'  \yyyy    = 4 digit year
'  \mm      = 2 digit month
'  \dd      = 2 digit day
'  \hh      = 2 digit hour
'  \mn      = 2 digit minute
'  \ss      = 2 digit second
'  \ts      = 2 digit test status (Reporter.RunStatus & Err.Number)
'  \level   = string of message level, eg Fatal, Debug, etc.
'  \indent  = variable number of spaces, used for doing in log call tracing
'  \routine = name of the calling routine
'  \message = the message to post to the appender
'
'-------------------------------------------------------------------------------
Function ClassAppender_Format(self,args)
   BagDictCreate args, NULL 

   ' args keys:
   '  message - the message object
   '  format  - the string

   Dim result
   ' set a default format
   result = "\yyyy\mm\dd \hh\mn\ss \ts \level \indent \routine \message"
   ' see if there's an explicit format coming from the appender write routine
   ' this would be the defaults defined in this file. usually, there will be
   If args.Exists("format") Then
      result = LCase(args("format"))
   End If
   ' see if the appender has an internally defined format
   ' this would have been specified at appender instantiation
   ' and so would indicate that the caller wanted to override 
   ' the defaults defined here
   If self.Exists("format") Then
      result = LCase(self("format"))
   End If
   ' finally see if the calling routine passed in an overriding format
   If args("message").Exists("format") Then
      result = LCase(args("message")("format"))
   End If

   ' now do the string replacements for the formatting
   result = Replace(result,"\yyyy",DatePart("YYYY",Now))
   result = Replace(result,"\yy", Right(DatePart("YYYY",Now),2))
   result = Replace(result,"\mm", Right("0" & DatePart("M",Now),2))
   result = Replace(result,"\dd", Right("0" & DatePart("D",Now),2))
   result = Replace(result,"\hh", Right("0" & Hour(Now),2))
   result = Replace(result,"\mn", Right("0" & Minute(Now),2))
   result = Replace(result,"\ss", Right("0" & Second(Now),2))
   result = Replace(result,"\ts", Reporter.RunStatus & Err.Number)

   ' and sort out the levels
   Dim loggingString
   loggingString = "LOGGING_LEVEL_" & args("message")("level")
   loggingString = Replace(loggingString,"-","_")
   result = Replace(result,"\level", Left(Eval(loggingString) & String(LOGGING_LEVELS_WIDTH, " "),LOGGING_LEVELS_WIDTH) )

   ' as of Aug 26 2009 Indenting is not yet enabled. When it is, the empty strings
   ' in the code below will be replaced by calculated numbers of spaces in order
   ' to trace entry and exit from functions - indenting is useless without 
   ' LogTraceEnter and LogTraceExit
   If self("indent_enabled") = True Then
      'result = Replace(result,"\indent","-TBD:INDENT_NOT_IMPLEMENTED-")
      result = Replace(result,"\indent","")
   Else
      result = Replace(result,"\indent","")
   End If
   
   ' and lastly, we replace the routine name and message text
   result = Replace(result,"\routine", args("message")("routine"))
   result = Replace(result,"\message", args("message")("message"))
   
   ClassAppender_Format = result
End Function
classAppender.ApplyProp "Format", "Get", "vector=>ClassAppender_Format|return_type=>string"

'-------------------------------------------------------------------------------
' Bag Method: Add (ClassAppender_Add)
'
'  Adds self to the loggingConfiguration("appenders.queue")
'  This queue is crawled by the Message.Dispatch BagMethod
'
' Parameters:
'   - self("name") - name of this appender
'
' Returns:
'   - no return value
'
' Usage:
'  tempAppender.Method "Add",NULL
'
' Notes:
'  3 appenders are allowed to be redefined. This is because they are defined
'  by default. No other appenders are allowed to be redefined. The 3 that are 
'  allowed are named: fatal_stop, reporter_fatal and print
'
'  Pushes the current appender onto the ClassQueue in loggingConfiguration("appenders.queue")
'
'-------------------------------------------------------------------------------
Function ClassAppender_Add(self, args)
   ClassAppender_Add = NULL
   ' am i allowed to write this appender into the stack?
   Dim appendersAllowedToBeRedefined, allowToWrite
   appendersAllowedToBeRedefined = "|fatal_stop|reporter_fatal|print|"
   allowToWrite = NOT loggingConfiguration("appenders.queue").Exists (self("name"))
   If InStr(1,appendersAllowedToBeRedefined,"|" & self("name") & "|") Then
      allowToWrite = True
   End If

   ' either write it or error out
   If allowToWrite Then
      loggingConfiguration("appenders.queue").Method "Push",Array("name", self("name"), "item", self)
   Else
      print "TBD:FATAL: attempted to instantiate appender " & self("name") & ", which already exists (and is not a redefinable appender)."
      Err.Raise 1,"", "TBD:FATAL: attempted to instantiate appender " & self("name") & ", which already exists (and is not a redefinable appender)."
   End If
   
End Function
classAppender.ApplyMethod "Add", NULL ' adds appender to queue

'-------------------------------------------------------------------------------
' Bag Method: Close (ClassAppender_Close)
'
'  Closes the appender target, removes self from appender queue in loggingConfiguration("appenders.queue")
'
' Parameters:
'   - none
'
' Usage:
'  myAppenderInstance.Method "Close", NULL
'
'-------------------------------------------------------------------------------
Function ClassAppender_Close(self, args)

   If NOT IsEmpty(self("file_system_object")) Then
      self("file_system_object").Close
   End If
   loggingConfiguration("appenders.queue")("data").Remove self("name")
   self.Method "Destroy",NULL

End Function
classAppender.ApplyMethod "Close", NULL ' by default, removes the appender from the queue

'-------------------------------------------------------------------------------
' Bag Method: SimpleFatal (classAppender_SimpleFatal)
'
'  Takes a picture, checks args("message")("do_not_exit"), and optionally exits
'  if loggingConfiguration("execute_on_fatal") is populated, call those routines
'
' Parameters:
'   - args("message")("do_not_exit")           - tells method to exit or not
'   - loggingConfiguration("execute_on_fatal") - list of things to call if exiting
'
' Returns:
'   - loggingConfiguration("fatal_encountered") gets set to True
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "SimpleFatal", Array("message",mymessageObject)
'
' Notes:
'  Unlike all other appenders, this one can close the program. 
'  Because of this, it gets queued to be executed LAST
'
'-------------------------------------------------------------------------------
Function ClassAppender_SimpleFatal(self, args)

   Dim i, currentAppenderKey, currentAppender
   LogPicture NULL
   If NOT args("message")("do_not_exit") = True Then
   
      loggingConfiguration("fatal_encountered") = True
      
      If loggingConfiguration("execute_on_fatal").Count > 0 Then
      
         For i = loggingConfiguration("execute_on_fatal").Count-1 to 0 Step -1
         
            currentAppenderKey = loggingConfiguration("execute_on_fatal").KeyIndex(i)
            currentAppender = loggingConfiguration("execute_on_fatal")(currentAppenderKey)
            Execute currentAppender
            
         Next 
         
      End If

      LoggingClose
      ExitTest Reporter.RunStatus
      ExitAction Reporter.RunStatus
      ExitRun  
      
   End If
   
End Function
classAppender.ApplyMethod "SimpleFatal", NULL

'-------------------------------------------------------------------------------
' Bag Method: SimpleReporterFatal (ClassAppender_SimpleReporterFatal)
'
'  Uses the QTP Reporter to post a fatal error message
'
' Parameters:
'   - args("message")("routine") - passed directly to reporter
'   - args("message")("message") - passed directly to reporter
'
' Returns:
'   - no return value
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "SimpleReporterFatal", Array("message",mymessageObject)
'
'-------------------------------------------------------------------------------
Function ClassAppender_SimpleReporterFatal(self, args)
   Reporter.ReportEvent micFail, args("message")("routine"), args("message")("message")
End Function
classAppender.ApplyMethod "SimpleReporterFatal", NULL

'-------------------------------------------------------------------------------
' Bag Method: SimplePrint (ClassAppender_SimplePrint)
'
'  prints the message to the QTP console window
'  default format is "\ts \level \routine: \message"
'
' Parameters (required):
'   - args("message")("routine") - calling routine
'   - args("message")("message") - message to print
'
' Parameters (allowed):
'   - self("format")             - formatting string specified at appender instantiation
'   - args("message")("format")  - formatting specified by the caller
'
' Returns:
'   - no return value
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "SimplePrint", Array("message",mymessageObject)
'
'-------------------------------------------------------------------------------
Function ClassAppender_SimplePrint(self, args)
   'print Reporter.RunStatus & ":" & args("message")("routine") & ": " & args("message")("message")

   ' does the appender define a format?
   If self.Exists("format") Then
      args("format") = self("format")
   End If
   ' if none is still defined, use the default
   args.SetIfUndefined("format") = "\ts \level \routine: \message"  '\yyyy\mm\dd \hh\mn\ss \ts \level \indent \routine \message

   print self.Prop ("Format",args) ' args contains key:message
   args.Remove("format")
   
End Function
classAppender.ApplyMethod "SimplePrint", NULL

'-------------------------------------------------------------------------------
' Bag Method: SimpleUDP (ClassAppender_SimpleUDP)
'
'  sends the message to the UDP Logger
'  default format is "\yyyy\mm\dd \hh\mn\ss,\level,\ts,\indent\routine, \message"
'
' Parameters (required):
'   - args("message")("routine") - calling routine
'   - args("message")("message") - message to print
'   - self("file_name")          - the target file name (typically added at instantiation)
'   - self("host")               - the target host (typically added at instantiation)
'   - self("port")               - the port number  (typically added at instantiation)
'
' Parameters (allowed):
'   - self("format")             - formatting string specified at appender instantiation
'   - args("message")("format")  - formatting specified by the caller
'
' Returns:
'   - no return value
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "SimpleUDP", Array("message",mymessageObject)
'
' Notes:
'  Remember: The UDP logger must be running somewhere, and the instantiation 
'  must have specified the host and file details. 
'
'  This component is self initializing, and has a dependency on CreateObject("WinsckW.WinSock")
'
'-------------------------------------------------------------------------------
Function ClassAppender_SimpleUDP (self, args)

   If NOT (self("udp_is_initialized") = True) Then
      self.VerifyHasKeys Array("host","port","file_name"),"UDP Appender called without needed keys"

      Set self("log_socket") = CreateObject("WinsckW.WinSock")
      self("log_socket").Protocol = 1
      self("log_socket").RemoteHost = self("host")
      self("log_socket").RemotePort = self("port")

      self("udp_is_initialized") = True
   End If
   
   Dim localmessage
   ' does the appender define a format?
   If self.Exists("format") Then
      args("format") = self("format")
   End If
   ' if none is still defined, use the default
   args.SetIfUndefined("format") = "\yyyy\mm\dd \hh\mn\ss,\level,\ts,\indent\routine, \message"  '\yyyy\mm\dd \hh\mn\ss \ts \level \indent \routine \message
   ' net get the message thru the format
   localmessage = self.Prop ("Format",args)

   self("log_socket").SendData self("file_name") & "##" & localmessage ' timeStamp & "; " & level & "; " & message

End Function
classAppender.ApplyMethod "SimpleUDP", NULL

'-------------------------------------------------------------------------------
' Bag Method: SimpleFile (ClassAppender_SimpleFile)
'
'  sends the message to the specified file in APPEND mode
'  default format is "\yyyy\mm\dd \hh\mn\ss,\ts,\level,\indent\routine,\message"
'
' Parameters (required):
'   - args("message")("routine") - calling routine
'   - args("message")("message") - message to print
'   - self("file_name")          - the target file name (typically added at instantiation)
'
' Parameters (allowed):
'   - self("format")             - formatting string specified at appender instantiation
'   - args("message")("format")  - formatting specified by the caller
'
' Returns:
'   - no return value
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "SimpleFile", Array("message",mymessageObject)
'
' Notes:
'  This component is self initializing
'
'-------------------------------------------------------------------------------
Function ClassAppender_SimpleFile(self,args)

   ' does the appender define a format?
   If self.Exists("format") Then
      args("format") = self("format")
   End If
   ' if none is still defined, use the default
   args.SetIfUndefined("format") = "\yyyy\mm\dd \hh\mn\ss,\ts,\level,\indent\routine,\message"  '\yyyy\mm\dd \hh\mn\ss \ts \level \indent \routine \message
   ' net get the message thru the format

   Dim textToPostToFile, fileName
   fileName = self("file_name")
   textToPostToFile = self.Prop ("Format",args)

   If IsEmpty(self("file_system_object")) Then
		Set self("file_system_object") = ClassFileSystemItem_FSO.OpenTextFile(fileName, LOG_FileIO_ForAppending, True)
	End If
   
   WriteToFile self.localData("file_system_object"), fileName, textToPostToFile
      
End Function
classAppender.ApplyMethod "SimpleFile", NULL

'-------------------------------------------------------------------------------
' Function: WriteToFile 
'
'  Helper function. Appends the passed data to the file. Opens the file as needed.
'  Used in multiple places in the logging code.
'
' Parameters:
'   - fileObject       - The file object from the OpenTextFile call. Can be uninitialized
'   - fileName         - the name of the file to write to
'   - textToPostToFile - the text to post into the file
'
' Returns:
'   - no return value
'
' Usage:
'  WriteToFile myFileObject, "foo.txt", "some text to post here"
'
' Notes:
'  This code appends a vbCrLf to the end of the passed string
'
'-------------------------------------------------------------------------------
Function WriteToFile(byRef fileObject, fileName, textToPostToFile)

   If IsNull(fileObject) OR IsEmpty(fileObject) Then
      Set fileObject = ClassFileSystemItem_FSO.OpenTextFile(fileName, LOG_FileIO_ForAppending, True)
   End If
   
   On Error Resume Next
	fileObject.Write( textToPostToFile & vbCrLf )
	On Error Goto 0

   If Err.Number > 0 Then
      fileObject.Close
		Err.Clear
      Set fileObject = ClassFileSystemItem_FSO.OpenTextFile(fileName, LOG_FileIO_ForAppending, True)
      fileObject.Write( textToPostToFile & vbCrLf )
	End If

End Function

'-------------------------------------------------------------------------------
' Bag Method: SimpleTestResult (ClassAppender_SimpleTestResult)
'
'  sends the message to the specified file in APPEND mode
'  default format is "\yyyy\mm\dd \hh\mn\ss,\ts,\message"
'
' Parameters (required):
'   - args("message")("routine") - calling routine
'   - args("message")("message") - message to print
'   - self("file_name")          - the target file name (typically added at instantiation)
'
' Parameters (allowed):
'   - self("format")             - formatting string specified at appender instantiation
'   - args("message")("format")  - formatting specified by the caller
'
' Returns:
'   - no return value
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "SimpleTestResult", Array("message",mymessageObject)
'
' Notes:
'  This component is self initializing
'
'-------------------------------------------------------------------------------
Function ClassAppender_SimpleTestResult (self, args)

   ' does the appender define a format?
   If self.Exists("format") Then
      args("format") = self("format")
   End If
   ' if none is still defined, use the default
   args.SetIfUndefined("format") = "\yyyy\mm\dd \hh\mn\ss,\ts,\message"  '\yyyy\mm\dd \hh\mn\ss \ts \level \indent \routine \message
   ' net get the message thru the format

   Dim textToPostToFile, fileName, TESTME
   fileName = self("file_name")
   If IsEmpty(fileName) Then
      print "TBD:FATAL: attempted to write to logging file via ClassAppender_SimpleTestResult without a file name defined"
      Err.Raise 1,"", "TBD:FATAL: attempted to write to logging file via ClassAppender_SimpleTestResult without a file name defined"
   End If

   textToPostToFile = self.Prop ("Format",args)

   If IsEmpty(self("file_system_object")) Then
      Set TESTME = ClassFileSystemItem_FSO.OpenTextFile(fileName, LOG_FileIO_ForAppending, True)
		Set self("file_system_object") = TESTME
	End If
   
   WriteToFile self.localData("file_system_object"), fileName, textToPostToFile
   
End Function
classAppender.ApplyMethod "SimpleTestResult", NULL

'-------------------------------------------------------------------------------
' Bag Method: SimplePicture (ClassAppender_SimplePicture)
'
'  snaps a picture of the screen, saves it to file
'  file name is built at runtime from the passed file name and an index
'  the file name should contain the string: "<index>" (yes, with the angle brackets)
'  a debug message is also written to the log
'
' Parameters (required):
'   - args("message")("routine") - calling routine
'   - args("message")("message") - message to print
'   - self("file_name")          - the target file name (typically added at instantiation)
'
' Returns:
'   - no return value
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "SimplePicture", Array("message",mymessageObject)
'
'-------------------------------------------------------------------------------
Function ClassAppender_SimplePicture (self, args)

   Dim localMessage, localRoutine
   Dim fileName

   localRoutine = "ClassAppender_SimplePicture"
   If args("message").Exist("routine") Then
      If NOT args("message")("routine") = "WAS_NOT_DEFINED" Then
         localRoutine = args("message")("routine")
      End If      
   End If

   If args("message").Exist("message") Then
      If NOT args("message")("message") = "WAS_NOT_DEFINED" Then
         localMessage = "routine=>" & localRoutine & "|message=>" & args("message")("message")
         LogInfo localMessage
      End If      
   End If

   fileName = self("file_name")

   self.SetIfUndefined("index") = 0
   
   fileName = Replace(fileName, "<index>", self("index"))
   self("index") = self("index") + 1

   LogDebug "routine=>" & localRoutine & "|message=>Saving screen shot " & fileName   
   Desktop.CaptureBitmap fileName, True
   
End Function
classAppender.ApplyMethod "SimplePicture", NULL



'============================================================================================================
'-------------------------------------------------------------------------------
' loggingConfiguration is basically all the global logging info... appenders queue,
'                      stack to execute on fatal, etc.
'-------------------------------------------------------------------------------
Dim loggingConfiguration
Set loggingConfiguration = BagDictMake(NULL,NULL)


'-------------------------------------------------------------------------------
' Function: LoggingInitialize
'
'  Performs default initialization of logging functions. Sets up queues, print,
'  fatal_stop, and reporter_fatal error handlers.
'
' Parameters:
'   - None
'
' Returns:
'   - No return value
'
' Notes:
'  Executed automatically inline, no need to run this explicitly within QTP
'
'-------------------------------------------------------------------------------
Function LoggingInitialize

   ' functional initialization
   Set loggingConfiguration("appenders.queue") = classQueue.NewObject(NULL,NULL)
   Set loggingConfiguration("execute_on_fatal") = BagDictMake(NULL,NULL)
   Set loggingConfiguration("call_stack.level_6") = classQueue.NewObject(NULL,NULL)
   loggingConfiguration("call_stack.level_6.id") = 1
   loggingConfiguration("initialization.successful") = False

   ' default appenders
   Dim tempAppender
   ' appender fatal: the first one into the queue must be the fatal one, the last one to be executed 
   Set tempAppender = classAppender.NewObject(NULL,NULL)
   tempAppender.ApplyKeys "name=>fatal_stop"
   tempAppender.ApplyKeys "level=><eval>LOGGING_LEVEL_FAIL" ' test failures are fatal errors by default
   tempAppender.ApplyMethod "Write", "vector=>ClassAppender_SimpleFatal"
   tempAppender.Method "Add",NULL

   ' appender print: the default 'get it to the console' print handler. this default one
   ' sets the level to catch everything. more usual for us is LOGGING_LEVEL_INFO
   Set tempAppender = classAppender.NewObject(NULL,NULL)
   tempAppender.ApplyKeys "name=>print"                        ' print is instantiated after reporter so reporter results can be shown in print
   tempAppender.ApplyKeys "level=><eval>LOGGING_LEVEL_SPECIAL" ' usual is LOGGING_LEVEL_INFO"
   tempAppender.ApplyMethod "Write", "vector=>ClassAppender_SimplePrint"
   tempAppender.Method "Add",NULL

   ' appender reporter_fatal: sets the message in the qtp reporter object
   Set tempAppender = classAppender.NewObject(NULL,NULL)
   tempAppender.ApplyKeys "name=>reporter_fatal"
   tempAppender.ApplyKeys "level=><eval>LOGGING_LEVEL_FAIL"
   tempAppender.ApplyMethod "Write", "vector=>ClassAppender_SimpleReporterFatal"
   tempAppender.Method "Add",NULL

   ' complete
   loggingConfiguration("fatal_encountered") = False
   loggingConfiguration("initialization.successful") = True

End Function
LoggingInitialize

'-------------------------------------------------------------------------------
' Function: LoggingClose
'
'  Closes all the appenders, removing them from the queues
'
' Parameters:
'   - none
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
function LoggingClose

   Dim i, currentAppenderKey, currentAppender
   If loggingConfiguration("appenders.queue")("data").Count > 0 Then
   
      For i = loggingConfiguration("appenders.queue")("data").Count-1 to 0 Step -1 ' this gives us the pointer into the array

         currentAppenderKey = loggingConfiguration("appenders.queue")("data").KeyIndex(i)
         Set currentAppender = loggingConfiguration("appenders.queue")("data")(currentAppenderKey)
         currentAppender.Method "Close",NULL
         
      Next

   End If
      
End function

'-------------------------------------------------------------------------------
' Function: LoggingInternalCommon
'
'  Used interally in logging, called by most public interfaces to perform 
'  the dispatch
'
' Parameters:
'   - args               - BagDict - The message
'   - localArgs          - BagDict compatible arguments
'   - localArgs("level") - the level of the error (by default supplied by the 
'                          public interface)
'   - localArgs("tags")  - tags from the public interface
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Private Function LoggingInternalCommon (args, localArgs)
   Dim theMessage, foo, rawArgs, passedLevel
   Assign rawArgs, args
   
   DictMake localArgs, NULL
   passedLevel = localArgs("level")

   If loggingConfiguration("initialization.successful") = True Then
      Set theMessage = classMessage.NewObject(args,NULL)
      theMessage.SetIfUndefined("level") = passedLevel
      If theMessage("message") = "WAS_NOT_DEFINED" Then
         If IsString(rawArgs) Then
            theMessage("message") = rawArgs
         End If
      End If
      Set theMessage("public_interface_data") = localArgs
      theMessage.Method "Dispatch",NULL
   Else
      On Error Resume Next
      Set theMessage = DictCreate(args)
      If IsString(rawArgs) AND IsEmpty(theMessage("message")) Then
         theMessage("message") = rawArgs
      End If
      If IsEmpty(theMessage("routine")) Then
         theMessage("routine") = "Routine WAS_NOT_DEFINED"
      End If
      Reporter.ReportEvent micFail, theMessage("routine"), theMessage("message")
      print "FATAL ERROR BEFORE LOGGING INITIALIZATION COMPLETE"
      print Reporter.RunStatus & ":" & theMessage("routine") & ": " & theMessage("message")
      Err.Raise 1,"", "FATAL ERROR BEFORE LOGGING INITIALIZATION COMPLETE"
      foo = MsgBox("FATAL ERROR BEFORE LOGGING INITIALIZATION COMPLETE", vbSystemModal AND vbCritical, "FATAL ERROR")
      ExitTest
      On Error Goto 0
   End If
   
   If loggingConfiguration("fatal_encountered") = True Then
      ExitTest Reporter.RunStatus
      ExitAction Reporter.RunStatus
      ExitRun   
   End If
   
End function

'-------------------------------------------------------------------------------
' Function: TagMerge
'
'  Returns a string with merged tags. does not copy tags that already exist
'  in the target.
'
' Parameters:
'   - oldTags - the space seperated tag string to append new tags to
'   - newTags - the space seperated tag string to copy new tags from
'
' Returns:
'  - string of the whole set of tags
'
'-------------------------------------------------------------------------------
Function TagMerge(byRef oldTags, newTags)
   Dim result, numItemsFound, storedItemsArray, i
   result = ""
   storedItemsArray = Split(oldTags & " " & newTags, " ")
   For i = 0 to UBound(storedItemsArray)
      If InStr(1,result, storedItemsArray(i)) = 0 Then
         result = result & " " & storedItemsArray(i)
      End If
   Next
   TagMerge = result
End Function

'-------------------------------------------------------------------------------
' Function: LogFatal
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  Logs a fatal issue internal to the automation, not related to a test failure
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogFatal(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_FATAL, "tags", "architect.fatal developer.fatal developer.fail step.fail")
End Function

'-------------------------------------------------------------------------------
' Function: LogFail
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  This interface logs a test failure
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogFail(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_FAIL, "tags", "developer.fail step.fail")
End Function

'-------------------------------------------------------------------------------
' Function: LogFailure
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  This interface logs a test failure
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogFailure(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_FAIL, "tags", "developer.fail step.fail")
End Function

'-------------------------------------------------------------------------------
' Function: LogWarning
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogWarning(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_WARN, "tags", "developer.warn")
End Function

'-------------------------------------------------------------------------------
' Function: LogWarn
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogWarn(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_WARN, "tags", "developer.warn")
End Function

'-------------------------------------------------------------------------------
' Function: LogInfo
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  Logs notes for the person evaluating the test results. Might indicate
'  for instance that the automation is about to try something (so that if
'  there's a failure, it becomes clear what the automation was trying to do)
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogInfo(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_INFO, "tags", "session.info test.info step.info")
End Function

'-------------------------------------------------------------------------------
' Function: LogPrint
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  Logs notes for the person evaluating the test results. Might indicate
'  for instance that the automation is about to try something (so that if
'  there's a failure, it becomes clear what the automation was trying to do)
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogPrint(args) 
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_INFO, "tags", "session.info test.info step.info")
End Function

'-------------------------------------------------------------------------------
' Function: LogAction
'
'  Used for logging UI or API Actions
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  For posting messages to the test developer
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogAction(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_ACTION)
End Function



'-------------------------------------------------------------------------------
' Function: LogDebug
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  For posting messages to the test developer
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogDebug(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_DEBUG, "tags", "developer.debug")
End Function

'-------------------------------------------------------------------------------
' Function: LogStatus
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  For posting messages to the Framework developer
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogStatus(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_STATUS, "tags", "architect.debug")
End Function

'-------------------------------------------------------------------------------
' Function: LogTrace
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  For posting messages to the Framework developer, part of the call stack system
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogTrace(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_TRACE, "tags", "architect.debug")
End Function

'-------------------------------------------------------------------------------
' Function: LogPicture
'
'  Logging public interface. Applies appropriate level and call the
'  common internal interface.
'
'  Snaps a picture of the screen and saves it to disk. notes this action
'  in the logs as a debug message
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogPicture(args)
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_PICTURE, "tags", "step.info")
End Function

'-------------------------------------------------------------------------------
' Function: LogResult
'
'  Notes a final test result at the end of a test. Usually posted to a CSV file
'  in the form: "time, 0, Pass, testcase name"
'
' Parameters:
'   - args - the message data
'
' Returns:
'   - No return value
'
'-------------------------------------------------------------------------------
Function LogResult(args)
   BagDictCreate args, NULL
	dim status
	status = "Pass"
	If Reporter.RunStatus = 1 Then ' 1 = fail, 3 = warn... 2009/11/04 removed "or Reporter.RunStatus = 3"
			status = "Fail"
	End If
	status = status & ", " & Environment("TestDir")
   If args.Exist("message") Then
      args("message") = status & "," & args("message")
   Else
      args("message") = status 
   End If
   LoggingInternalCommon args, Array("level",LOGGING_LEVEL_RESULT, "tags", Iif(status = "Pass","step.pass","step.fail"))
End Function

'-------------------------------------------------------------------------------
' Function: LogTraceEnter
'
'  These calls posts an entry to the debugging logs for each trace enter and exit
'
'  Because VBScript lacks a call stack, this provides that functionality
'  LogTraceEnter returns a unique object that is used at the end of the function
'  to call LogTraceExit.
'
' Parameters:
'   - args("routine") - the routine that's just been entered
'   - args("message") - any message you want to add
'   - passedArgs      - arguments passed to the calling function 
'
' Returns:
'   - a call stack object pointer
'
' Usage:
'  MyFunction(foo, bar)
'     Dim tracePointer
'     Set tracePointer = LogTraceEnter("routine=>MyFunction",Array("foo",foo,"bar",bar))
'     ...
'     MyFunction = resultOfMyFunctionCode
'     LogTraceExit tracePointer, MyFunction, NULL
'  End Function
'
' Notes:
'  This was necessary because VBScript's debugger does not include a callstack
'  function. This pair of functions not only maintains a call stack, but it also
'  reports an error when a function thusly insturmented terminates abnormally,
'  leaving it's object on the call stack. This can easily happen invisibly when
'  the call to the function happens deep inside an On Error Resume Next (calls 
'  making subcalls making subcalls without appropriate trapping).
'-------------------------------------------------------------------------------
Function LogTraceEnter(args,passedArgs)
   BagDictCreate args, NULL
   BagDictCreate passedArgs, NULL

   Dim callStackItem
   BagDictCreate callStackItem, NULL

   callStackItem("id") = loggingConfiguration("call_stack.level_6.id")
   loggingConfiguration("call_stack.level_6.id") = loggingConfiguration("call_stack.level_6.id") + 1
   
   Set callStackItem("parameters.raw") = passedArgs
   callStackItem("parameters.string") = passedArgs.RenderAsString()
   callStackItem("entry_time") = Now
   callStackItem("routine") = args("routine")
   callStackItem("message") = args("message")
   'callStackItem("") = args("")
   
   loggingConfiguration("call_stack.level_6").Method "Push",Array("name", CString(callStackItem("id")), "item", callStackItem)

   LogTrace Array("routine",args("routine"),_
                  "message","ENTER " & args("routine") & "(" & callStackItem("parameters.string") & ")")
   'GlobalDictionary("logging.appender.indentsize") = GlobalDictionary("logging.appender.indentsize") + 3
   
	set LogTraceEnter = callStackItem

End Function

'-------------------------------------------------------------------------------
' Function: LogTraceExit
'
'  Complimentary companion to LogTraceEnter, recieves the stack pointer, 
'  the function return value (or null), Verifies stack integrity, logs the event, 
'  and removes the item from the stack.
'
'  Because VBScript lacks a call stack, this provides that functionality
'  LogTraceEnter returns a unique object that is used at the end of the function
'  to call LogTraceExit.
'
' Parameters:
'   - stackPointer    - the object returned by LogTraceEnter
'   - resultValue     - the return value of the calling function, or NULL
'   - args            - future expansion
'
' Returns:
'   - no return value
'
' Usage:
'  MyFunction(foo, bar)
'     Dim tracePointer
'     Set tracePointer = LogTraceEnter("routine=>MyFunction",Array("foo",foo,"bar",bar))
'     ...
'     MyFunction = resultOfMyFunctionCode
'     LogTraceExit tracePointer, MyFunction, NULL
'  End Function
'
' Notes:
'  This was necessary because VBScript's debugger does not include a callstack
'  function. This pair of functions not only maintains a call stack, but it also
'  reports an error when a function thusly insturmented terminates abnormally,
'  leaving it's object on the call stack. This can easily happen invisibly when
'  the call to the function happens deep inside an On Error Resume Next (calls 
'  making subcalls making subcalls without appropriate trapping).
'-------------------------------------------------------------------------------
Function LogTraceExit(stackPointer, resultValue, args) 

   Dim temp
   ' is the passed in item the top of the stack?
   If stackPointer is loggingConfiguration("call_stack.level_6")("data").ItemIndex(-1) Then
      ' all is well, log the fact, pop the item, and be done
      LogTrace Array("routine",stackPointer("routine"),_
                     "message","EXIT " & stackPointer("routine") & " = " & CString(resultValue))
      Set temp = loggingConfiguration("call_stack.level_6").Prop("Pop",NULL)
   Else
      ' all is NOT WELL
      LogTrace Array("routine",stackPointer("routine"),_
                     "message","STACK CORRUPT - SOMETHING FAILED TO COMPLETE NORMALLY " & stackPointer("routine") & " = " & CString(resultValue))

      ' is the specified item still on the stack?
      If loggingConfiguration("call_stack.level_6")(loggingConfiguration("call_stack.level_6")("id")) Then
         ' this is the most likely case, that something this routine called 
         ' failed and left debris on the stack
         LogWarning "routine=>LogTraceExit|message=>Called with leftover items on the stack"
         Dim keepGoing, currentItem
         keepGoing = TRUE
         While keepGoing
            Set currentItem = loggingConfiguration("call_stack.level_6").Prop("Pop",NULL)
            If currentItem Is stackPointer Then
               keepGoing = False
            End If
            LogWarning "routine=>LogTraceExit|message=>Found entry for " & currentItem("routine") & "(" & currentItem("parameters.string") & ") (ID=" & currentItem("id") & ")"
         Wend
         LogFatal "routine=>LogTraceExit|message=>Called with left over items on stack, program is now unstable, stopping"
      Else
         ' this would be the most wacky case, where this ending point in the routine
         ' got called twice
         LogFatal "routine=>LogTraceExit|message=>Called without matching entry for " & stackPointer("routine") & "(" & stackPointer("parameters.string") & ") (ID=" & stackPointer("id") & ")"
      End If
   End If
   
End Function





'-------------------------------------------------------------------------------
' Bag Class: classAppenderForDatabase
' 
'  Writes messages to the logging database object
'
' Inherits: 
'  ClassAppender
'
' Implements:
'
' Notes:
'  as of this writing, the context list is located at:
'  Q:\Projects\Cross Project Documentation\Centralized Logging Design Notes - Tags.txt
'-------------------------------------------------------------------------------
Dim classTestContextMonitor
Set ClassTestContextMonitor = ClassBase.MakeClass("ClassTestContextMonitor", NULL, "self.is_virtual=><eval>False")
ClassTestContextMonitor("current_context") = "global"

Function ClassTestContextMonitor_Context_Get(self, args)
   ClassTestContextMonitor_Context_Get = self("current_context")
End Function
ClassTestContextMonitor.ApplyProp "context","get","vector=>ClassTestContextMonitor_Context_Get"
Function ClassTestContextMonitor_Context_Set(self, args)
   self("current_context") = args("item")
End Function
ClassTestContextMonitor.ApplyProp "context","set","vector=>ClassTestContextMonitor_Context_Set"

Dim GlobalContextMonitor
Set GlobalContextMonitor = ClassTestContextMonitor.NewObject(NULL,NULL)

'-------------------------------------------------------------------------------
' Bag Class: classAppenderForDatabase
' 
'  Writes messages to the logging database object
'
' Inherits: 
'  ClassAppender
'
' Implements:
'
'-------------------------------------------------------------------------------
Dim classAppenderForDatabase
Set ClassAppenderForDatabase = ClassAppender.MakeClass("classAppenderForDatabase", NULL, NULL)
ClassAppenderForDatabase.ApplyKeys     "level=>0"

'-------------------------------------------------------------------------------
' Bag Method: DoYouWantMe (classAppenderForDatabase_DoYouWantMe)
'
'  Overrides classAppender version of this call
'
' Parameters:
'   - no keys used
'
' Returns:
'  True (because this appender ALWAYS wants ALL the messages)
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: x = myAppenderInstance.Method "DoYouWantMe", Array("message",mymessageObject)
'
' Notes:
'
'-------------------------------------------------------------------------------
Function ClassAppenderForDatabase_DoYouWantMe(self, args)
   ' here we just always say YES
   ClassAppenderForDatabase_DoYouWantMe = self("enabled")
End Function
ClassAppenderForDatabase.ApplyMethod "DoYouWantMe", NULL

'-------------------------------------------------------------------------------
' Bag Method: Write (ClassAppenderForDatabase_Write)
'
'  Overrrides the inherited write method, adds logic to deal with database
'
' Parameters:
'   - key:args("message") - the message object to post
'   - key:args("message")("level") - used to determine tags
'
' Usage:
'  Typical usage has this called from ClassMessage.Dispatch, not directly
'  as shown here: myAppenderInstance.Method "Write", Array("message",mymessageObject)
'
' Notes: 
'
'-------------------------------------------------------------------------------
' some things we need: 
'  RemoteLogService
'  RemoteLogService.PostMessage (scriptID, sessionID, tags, message)
'  RemoteLogService.PostMessageEx (scriptID, sessionID, tags, message, attachments)
'   - scriptID: (string) the script ID
'   - sessionID: (string) the session ID
'   - tags: (string) the tags for the log event
'   - message: (string) the actual log message itself
'   - attachments: (string) a vbCR seperated list of the attachments we're
'                           including.  Each entry must be the full file path
' TestCase("name") TestCase("step")
Function ClassAppenderForDatabase_Write(self, args)
   Dim localTags, s, hasAttachmentsFlag, localSessionID, localTestID, localAttachments
   s = " " ' space, used for appending tags together
   localTags = ""
   hasAttachmentsFlag = False
   localSessionID = FrameworkConfiguration("session_id")
   localTestID = TestCase("name")
   localAttachments = args("attachments")
   ' CAN HAS LOGIC HERE FOR OVERRIDING WITH THE TC NAME FROM TEST RUNNER
   
   ' tag stamping rules
   On Error Resume Next
   localTags = TagMerge(localTags, args("tags"))
   localTags = TagMerge(localTags, args("message")("tags"))
   localTags = TagMerge(localTags, args("public_interface_data")("tags"))
   localTags = TagMerge(localTags, args("message")("public_interface_data")("tags"))
   On Error Goto 0
   
   ' now to post it
   Dim localPostResult
   If hasAttachmentsFlag Then
      localPostResult = RemoteLogService.PostMessageEx (localTestID, localSessionID, localTags, args("message")("message"), localAttachments)
   Else
      localPostResult = RemoteLogService.PostMessage (localTestID, localSessionID, localTags, args("message")("message"))
   End If
   If localPostResult = False Then
      self("enabled") = False ' we do this to prevent error recusion when we report the error
      On Error Resume Next
      LogWarning "routine=>ClassAppenderForDatabase_Write|message=>Error posting log entry to database: " & RemoteLogService.statusCode & " " & RemoteLogService.status
      LogWarning "routine=>ClassAppenderForDatabase_Write|message=>XML: " & vbCrLf & RemoteLogService.XML
      LogWarning "routine=>ClassAppenderForDatabase_Write|message=>Passed values: " & localTestID & ", " & localSessionID & ", " & localTags & ", " & args("message")("message")
      LogWarning "routine=>ClassAppenderForDatabase_Write|message=>args: " & vbCrLf & args.Debug()
      LogWarning "routine=>ClassAppenderForDatabase_Write|message=>args(""message""): " & vbCrLf & args("message").Debug()
      LogWarning "routine=>ClassAppenderForDatabase_Write|message=>args(""message"")(""public_interface_data""): " & vbCrLf & args("message")("public_interface_data").Debug()
      On Error Goto 0
      self("enabled") = True 
   End If
   
End Function
ClassAppenderForDatabase.ApplyMethod "Write", NULL ' by default vectors to simple print




'***********************************************************
' aggregating q:\utils\libs\class-centrallogmanager.vbs

'###############################################################################
' Library: CentralLogManager
'
' About: Description
'  This library provides a mechanism for logging events to a central remote server.
'
' Usage:
' The library already exports an instantiated object for you, named RemoteLogService
'
'###############################################################################
'Option Explicit

'===============================================================================
' Section: Constants and Globals
'===============================================================================

'-------------------------------------------------------------------------------
' Constants: Private Constants
'  Constants internal to the library
'
'   CENTRAL_LOG_SERVER  - The IP or hostname of the server that hosts the logging service
'   CENTRAL_LOG_SCRIPT  - The URL path on the data server to the logging service interface
'   CENTRAL_LOG_ATTACHMENT_PATH - The path to the SMB share where attachments are stored
'   CENTRAL_LOG_RETRIES - The number of times to retry the log post before failing
'-------------------------------------------------------------------------------
Private Const CENTRAL_LOG_SERVER = "10.4.64.132"
Private Const CENTRAL_LOG_SCRIPT  = "/qtplog/insert.cgi"
Private Const CENTRAL_LOG_ATTACHMENT_PATH = "\\10.4.64.132\QTP_LOG_ATTACHMENTS\" 
Private Const CENTRAL_LOG_RETRIES = 3

'-------------------------------------------------------------------------------
' Vars: Public Variables
'   Variables exported by the library
'
' RemoteLogService - A global instance of the CentralLogManager class, accessible from
'             any script that imports this library
'-------------------------------------------------------------------------------
Public RemoteLogService

'===============================================================================
' Class: CentralLogManager
'   The CentralLogManager class
'===============================================================================
Class CentralLogManager
   '===============================================================================
   ' Group: Private Members
   '   Class members that are internal and not available outside the class
   '===============================================================================

   '-------------------------------------------------------------------------------
   ' Vars: Private Data
   '   Private class data
   '
   ' xmlhttp - The  XMLHTTP object used to send the requests
   ' LastStatusCode - The HTTP status value returned from the last attempt to post the
   '              log message
   ' LastStatus - The status message returned from the last attempt to post the
   '              log message
   ' LastXML - The last log message (in XML format) that was sent, useful for debugging
   ' IsPTA   - True if the current user is a member of the test automation team, false
   '           if not.  If true, all messages will have a pta tag prepended to their tag list
   '-------------------------------------------------------------------------------
   Private xmlhttp
   Private LastStatusCode
   Private LastStatus
   Private LastXML
   Private IsPTA
   
   '-------------------------------------------------------------------------------
   ' Methods: Private Methods
   '-------------------------------------------------------------------------------

   '-------------------------------------------------------------------------------
   ' Sub: Class_Initialize
   '
   ' Constructor, initializes the class.  Creates the xmlhttp object
   '-------------------------------------------------------------------------------
   Private Sub Class_Initialize()
      Set xmlhttp = CreateObject("MSXML2.XMLHTTP.3.0")
      IsPTA = IsCurrentUserAMemberOfTheTestAutomationTeam
   End Sub

   '-------------------------------------------------------------------------------
   ' Sub: Class_Terminate
   '
   ' Destructor, finalizes the class.  Frees up the xmlhttp object
   '-------------------------------------------------------------------------------
   Private Sub Class_Terminate()
      Set xmlhttp = Nothing
   End Sub

   
   '-------------------------------------------------------------------------------
   ' Function: SaveAttachments
   '  Saves the attachments to the server.  Each attachment is saved with a unique name,
   ' made up of scriptID + sessionID + fileName.  The scriptID and sessionID are run 
   ' through Escape first, just to make sure there are no funny characters in there.
   ' The new file names are returned as a comma seperated list, ready for embedding in
   ' the log message.  NOTE, if any of the attachments can not be found, they are silently
   ' skipped.
   '
   ' Parameters:
   '   - scriptID: (string) the script ID
   '   - sessionID: (string) the session ID
   '   - attachments: (string) a comma seperated list of all the attachments we're
   '        including.
   '
   ' Returns:
   '  (string) the list of attachments as they are stored on the server
   '
   '------------------------------------------------------------------------------- 
   Private Function SaveAttachments (scriptID, sessionID, attachments)
      Dim prefix
      Dim fileList
      Dim buffer
      Dim newFileName
      Dim i
      
      prefix = scriptID & sessionID & "-"
      
      ' convert the string in to an array
      fileList = split(attachments, vbCr)
      
      for i = LBound(fileList) to UBound(fileList)
         if (FileExists(fileList(i))) then
            newFileName = prefix & FSO().GetFileName(fileList(i))
            FSO().CopyFile fileList(i), CENTRAL_LOG_ATTACHMENT_PATH & newFileName
            if (buffer = "") then
               buffer = newFileName
            else
               buffer = buffer & "|" & newFileName
            end if
         end if
      next
      
      SaveAttachments = buffer
   End Function
   
   '-------------------------------------------------------------------------------
   ' Function: BuildMessageBody
   '  Builds up the XML message body using the values passed in.  NOTE that this 
   ' function does no checking to make sure the values passed in are valid, that's
   ' done on the server side
   '
   ' Parameters:
   '   - scriptID: (string) the script ID
   '   - sessionID: (string) the session ID
   '   - tags: (string) the tags for this log event
   '   - message: (string) the actual log message itself
   '   - attachments: (string) a comma seperated list of all the attachments we're
   '        including (can be an empty string).  Files in this list should have already
   '        been processed by <SaveAttachments>
   '
   ' Returns:
   '  (string) the XML message body
   '
   '-------------------------------------------------------------------------------   
   Private Function BuildMessageBody (scriptID, sessionID, tags, message, attachments)
      Dim buffer
      
      if (IsPTA) then
         tags = "pta " & tags
      end if
      
      buffer = "<log_message>" & _
               "<script_id><![CDATA[" & scriptID & "]]></script_id>" & _
               "<session_id><![CDATA[" & sessionID & "]]></session_id>" & _
               "<tags><![CDATA[" & tags & "]]></tags>" & _
               "<message><![CDATA[" & message & "]]></message>" & _
               "<attachments><![CDATA[" & attachments & "]]></attachments>" & _
               "</log_message>"
      
      BuildMessageBody = buffer
   End Function

   '-------------------------------------------------------------------------------
   ' Function: ProcessHTTPRequest
   '  This sub uses the XMLHTTP object to actually post the log message to the
   '  server.  The status code returned from the server is stored into LastStatus and
   '  LastStatusCode.  The function will retry the post CENTRAL_LOG_RETRIES times if
   '  the post fails.
   '
   ' Parameters:
   '   - message: (string) the XML formatted log message to post
   '
   ' Returns:
   '  (boolean) true if the post succeeded, false if it didn't
   '-------------------------------------------------------------------------------
   Private Function ProcessHTTPRequest (message)
      Dim url
      Dim retries
      
      retries = 0
      ' Reset the stored status
      LastStatus = ""
      LastStatusCode = 0

      Do While (retries < CENTRAL_LOG_RETRIES and LastStatusCode <> 200)
         ' Use Timer to defeat caching
         url = "http://" & CENTRAL_LOG_SERVER & CENTRAL_LOG_SCRIPT & "?r=" & Timer
         LastXML = message
         
         'Make sure the third param to open is false so request is synchronous
         xmlhttp.open "POST", url, false
         xmlhttp.setRequestHeader "content-type", "text/xml"
         xmlhttp.setRequestHeader "content-length", len(message)
         xmlhttp.setRequestHeader "connection", "close"
         On Error Resume Next
         xmlhttp.send(message)
         If (Err) Then
            LastStatus = Err.Description
            LastStatusCode = Err.Number
            Err.Clear
            Exit Do
         End If
         On Error GoTo 0
         LastStatus = xmlhttp.statusText
         LastStatusCode = xmlhttp.status
         
         retries = retries + 1
      Loop

      
      if (LastStatusCode <> 200) then
         ProcessHTTPRequest = false
      else
         ProcessHTTPRequest = true
      end if
   End Function

   '===============================================================================
   ' Group: Public Members
   '   Class members that are available for use
   '===============================================================================

   '-------------------------------------------------------------------------------
   ' Properties: Public Properties
   '-------------------------------------------------------------------------------

   '-------------------------------------------------------------------------------
   ' Property: StatusCode (Get)
   '   Returns the current value stored in LastStatusCode
   '
   ' Returns:
   '  (integer) The value stored in LastStatusCode
   '-------------------------------------------------------------------------------
   Public Property Get StatusCode
      StatusCode = LastStatusCode
   End Property

   '-------------------------------------------------------------------------------
   ' Property: Status (Get)
   '   Returns the current value stored in LastStatus
   '
   ' Returns:
   '  (string) The value stored in LastStatus
   '-------------------------------------------------------------------------------
   Public Property Get Status
      Status = LastStatus
   End Property
   
   '-------------------------------------------------------------------------------
   ' Property: XML (Get)
   '   Returns the XML for the last log message that was posted
   '
   ' Returns:
   '  (string) The value stored in LastXML
   '-------------------------------------------------------------------------------
   Public Property Get XML
      XML = LastXML
   End Property

   '-------------------------------------------------------------------------------
   ' Methods: Public Methods
   '-------------------------------------------------------------------------------   
   
   '-------------------------------------------------------------------------------
   ' Function: PostMessage
   '  Posts a log message to the log server.  NOTE, this method works for log messages
   ' that have no attachments, if you wish to post attachments as well, use <PostMessageEx>
   '
   ' Parameters:
   '   - scriptID: (string) the script ID
   '   - sessionID: (string) the session ID
   '   - tags: (string) the tags for the log event
   '   - message: (string) the actual log message itself
   '
   ' Returns:
   '  (boolean) true if the log message was sent succesfully, false if it failed. If 
   '            the log message failed to send, check the values of the <StatusCode> and
   '            <Status> properties.
   '------------------------------------------------------------------------------- 
   Function PostMessage (scriptID, sessionID, tags, message)
      Dim xml
      xml = BuildMessageBody (scriptID, sessionID, tags, message, "")
      PostMessage = ProcessHTTPRequest(xml)
   End Function
   
   '-------------------------------------------------------------------------------
   ' Function: PostMessageEx
   '  Posts a log message to the log server.  NOTE, this method works for log messages
   ' that have attachments, if you aren't posting attachments use <PostMessage>
   '
   ' Parameters:
   '   - scriptID: (string) the script ID
   '   - sessionID: (string) the session ID
   '   - tags: (string) the tags for the log event
   '   - message: (string) the actual log message itself
   '   - attachments: (string) a vbCR seperated list of the attachments we're
   '                           including.  Each entry must be the full file path
   '
   ' Returns:
   '  (boolean) true if the log message was sent succesfully, false if it failed. If 
   '            the log message failed to send, check the values of the <StatusCode> and
   '            <Status> properties.
   '------------------------------------------------------------------------------- 
   Public Function PostMessageEx (scriptID, sessionID, tags, message, attachments)
      Dim xml
      Dim attachList
      
      ' save the attachments to the remote server and get a list of the unique file names on
      ' the remote server
      attachList = SaveAttachments(scriptID, sessionID, attachments)
      
      xml = BuildMessageBody (scriptID, sessionID, tags, message, attachList)
      PostMessageEx = ProcessHTTPRequest(xml)
   End Function
   
End Class

'===============================================================================
' Section: Public Functions
'   Functions that are exported by the library
'===============================================================================

'-------------------------------------------------------------------------------
' Function: FrameworkDetect_Class_CentralLogManager
'   Utility function for the Framework Compilation chechking utility
'
' Returns:
'  (integer) always returns 1
'
'-------------------------------------------------------------------------------
Public Function FrameworkDetect_Class_CentralLogManager()
	FrameworkDetect_Class_CentralLogManager = 1
End Function

'-------------------------------------------------------------------------------
' Function: NewCentralLogManager
'   This stupid function is needed because Mercury never saw fit to enable a user
'   to instantiate an instance of a class defined in an external library inside an
'   action (Mercury QTP KB Article 45025).  So we have to have this ugly kludge
'   to return us a new instance of the class
'
' Returns:
'  (object) an instance of CentralLogManager
'
'-------------------------------------------------------------------------------
Public Function NewCentralLogManager
    Set NewCentralLogManager = New CentralLogManager
End Function

' Create our published global instance of the CentralLogManager 
Set RemoteLogService = New CentralLogManager

'###############################################################################
' End Library PersistentDataStorage
'###############################################################################


