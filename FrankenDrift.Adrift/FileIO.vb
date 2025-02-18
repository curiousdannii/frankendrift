Imports System.IO
Imports System.Text.Encoding
Imports System.Xml

Imports Eto.Drawing
Imports FrankenDrift.Glue
Imports FrankenDrift.Glue.Util
Imports ICSharpCode.SharpZipLib
Imports ICSharpCode.SharpZipLib.Zip.Compression.Streams

Module FileIO

    Private br As BinaryReader
    Private bw As BinaryWriter
    Private sLoadString As String
    Private bAdventure() As Byte
    Private salWithStates As New StringArrayList
    Private iStartPriority As Integer
    Private dFileVersion As Double
    Friend Enum LoadWhatEnum
        All
        Properies
        AllExceptProperties
    End Enum


    Private Function LoadDescription(ByVal nodContainerXML As XmlElement, ByVal sNode As String) As Description

        Dim Description As New Description

        If nodContainerXML.Item(sNode) Is Nothing OrElse nodContainerXML.GetElementsByTagName(sNode).Count = 0 Then Return Description

        Dim nodDescription As XmlElement = CType(nodContainerXML.GetElementsByTagName(sNode)(0), XmlElement)

        For Each nodDesc As XmlElement In nodDescription.SelectNodes("Description") '. nodDescription.GetElementsByTagName("Description")
            Dim sd As New SingleDescription
            sd.Restrictions = LoadRestrictions(nodDesc)
            sd.eDisplayWhen = CType([Enum].Parse(GetType(SingleDescription.DisplayWhenEnum), nodDesc.Item("DisplayWhen").InnerText), SingleDescription.DisplayWhenEnum)
            If nodDesc.Item("Text") IsNot Nothing Then sd.Description = nodDesc.Item("Text").InnerText
            If nodDesc.Item("DisplayOnce") IsNot Nothing Then sd.DisplayOnce = GetBool(nodDesc.Item("DisplayOnce").InnerText)
            If nodDesc.Item("ReturnToDefault") IsNot Nothing Then sd.ReturnToDefault = GetBool(nodDesc.Item("ReturnToDefault").InnerText)
            If nodDesc.Item("TabLabel") IsNot Nothing Then sd.sTabLabel = SafeString(nodDesc.Item("TabLabel").InnerText)
            Description.Add(sd)
        Next

        If Description.Count > 1 Then Description.RemoveAt(0)
        Return Description

    End Function


    Private Sub SaveDescription(ByVal xmlWriter As XmlTextWriter, ByVal Description As Description, ByVal sNode As String)

        If Description.Count = 0 OrElse (Description.Count = 1 AndAlso Description(0).Description = "") Then Exit Sub

        With xmlWriter
            .WriteStartElement(sNode) ' Description
            For Each sd As SingleDescription In Description
                .WriteStartElement("Description") ' SingleDescription
                If sd.Restrictions.Count > 0 Then SaveRestrictions(xmlWriter, sd.Restrictions)
                .WriteElementString("DisplayWhen", sd.eDisplayWhen.ToString)
                .WriteElementString("Text", sd.Description)
                If sd.DisplayOnce Then .WriteElementString("DisplayOnce", "1")
                If sd.ReturnToDefault Then .WriteElementString("ReturnToDefault", "1")
                If sd.sTabLabel <> "Default Description" AndAlso Not sd.sTabLabel.StartsWith("Alternative Description ") AndAlso sd.sTabLabel <> "" Then
                    .WriteElementString("TabLabel", sd.sTabLabel)
                End If
                .WriteEndElement() ' Description
            Next
            .WriteEndElement() ' sNode
        End With
    End Sub


    ' Write a state to file
    Friend Function SaveState(ByVal state As clsGameState, ByVal sFilePath As String) As Boolean

        Try
            Dim stmMemory As New MemoryStream
            Dim xmlWriter As New System.Xml.XmlTextWriter(stmMemory, System.Text.Encoding.UTF8)
            Dim bData() As Byte

            With xmlWriter
                .Indentation = 4 ' Change later

                .WriteStartDocument()
                .WriteStartElement("Game")

                ' TODO: Ideally this should only save values that are different from the initial TAF state                

                For Each sKey As String In state.htblLocationStates.Keys
                    Dim locs As clsGameState.clsLocationState = CType(state.htblLocationStates(sKey), clsGameState.clsLocationState)
                    .WriteStartElement("Location")
                    .WriteElementString("Key", sKey)
                    For Each sPropKey As String In locs.htblProperties.Keys
                        Dim props As clsGameState.clsLocationState.clsStateProperty = CType(locs.htblProperties(sPropKey), clsGameState.clsLocationState.clsStateProperty)
                        .WriteStartElement("Property")
                        .WriteElementString("Key", sPropKey)
                        .WriteElementString("Value", props.Value)
                        .WriteEndElement()
                    Next
                    For Each sDisplayedKey As String In locs.htblDisplayedDescriptions.Keys
                        .WriteElementString("Displayed", sDisplayedKey)
                    Next
                    .WriteEndElement()
                Next

                For Each sKey As String In state.htblObjectStates.Keys
                    Dim obs As clsGameState.clsObjectState = CType(state.htblObjectStates(sKey), clsGameState.clsObjectState)
                    .WriteStartElement("Object")
                    .WriteElementString("Key", sKey)
                    If obs.Location.DynamicExistWhere <> clsObjectLocation.DynamicExistsWhereEnum.Hidden Then
                        .WriteElementString("DynamicExistWhere", obs.Location.DynamicExistWhere.ToString)
                    End If
                    If obs.Location.StaticExistWhere <> clsObjectLocation.StaticExistsWhereEnum.NoRooms Then
                        .WriteElementString("StaticExistWhere", obs.Location.StaticExistWhere.ToString)
                    End If
                    .WriteElementString("LocationKey", obs.Location.Key)
                    For Each sPropKey As String In obs.htblProperties.Keys
                        Dim props As clsGameState.clsObjectState.clsStateProperty = CType(obs.htblProperties(sPropKey), clsGameState.clsObjectState.clsStateProperty)
                        .WriteStartElement("Property")
                        .WriteElementString("Key", sPropKey)
                        .WriteElementString("Value", props.Value)
                        .WriteEndElement()
                    Next
                    For Each sDisplayedKey As String In obs.htblDisplayedDescriptions.Keys
                        .WriteElementString("Displayed", sDisplayedKey)
                    Next
                    .WriteEndElement()
                Next

                For Each sKey As String In state.htblTaskStates.Keys
                    Dim tas As clsGameState.clsTaskState = CType(state.htblTaskStates(sKey), clsGameState.clsTaskState)
                    .WriteStartElement("Task")
                    .WriteElementString("Key", sKey)
                    .WriteElementString("Completed", tas.Completed.ToString)
                    .WriteElementString("Scored", tas.Scored.ToString)
                    For Each sDisplayedKey As String In tas.htblDisplayedDescriptions.Keys
                        .WriteElementString("Displayed", sDisplayedKey)
                    Next
                    .WriteEndElement()
                Next

                For Each sKey As String In state.htblEventStates.Keys
                    Dim evs As clsGameState.clsEventState = CType(state.htblEventStates(sKey), clsGameState.clsEventState)
                    .WriteStartElement("Event")
                    .WriteElementString("Key", sKey)
                    .WriteElementString("Status", evs.Status.ToString)
                    .WriteElementString("Timer", evs.TimerToEndOfEvent.ToString)
                    .WriteElementString("SubEventTime", evs.iLastSubEventTime.ToString)
                    .WriteElementString("SubEventIndex", evs.iLastSubEventIndex.ToString)
                    For Each sDisplayedKey As String In evs.htblDisplayedDescriptions.Keys
                        .WriteElementString("Displayed", sDisplayedKey)
                    Next
                    .WriteEndElement()
                Next

                For Each sKey As String In state.htblCharacterStates.Keys
                    Dim chs As clsGameState.clsCharacterState = CType(state.htblCharacterStates(sKey), clsGameState.clsCharacterState)
                    .WriteStartElement("Character")
                    .WriteElementString("Key", sKey)
                    If chs.Location.ExistWhere <> clsCharacterLocation.ExistsWhereEnum.Hidden Then
                        .WriteElementString("ExistWhere", chs.Location.ExistWhere.ToString)
                    End If
                    If chs.Location.Position <> clsCharacterLocation.PositionEnum.Standing Then
                        .WriteElementString("Position", chs.Location.Position.ToString)
                    End If
                    If chs.Location.Key <> "" Then .WriteElementString("LocationKey", chs.Location.Key)
                    For Each ws As clsGameState.clsCharacterState.clsWalkState In chs.lWalks
                        .WriteStartElement("Walk")
                        .WriteElementString("Status", ws.Status.ToString)
                        .WriteElementString("Timer", ws.TimerToEndOfWalk.ToString)
                        .WriteEndElement()
                    Next
                    For Each sSeen As String In chs.lSeenKeys
                        .WriteElementString("Seen", sSeen)
                    Next
                    For Each sPropKey As String In chs.htblProperties.Keys
                        Dim props As clsGameState.clsCharacterState.clsStateProperty = CType(chs.htblProperties(sPropKey), clsGameState.clsCharacterState.clsStateProperty)
                        .WriteStartElement("Property")
                        .WriteElementString("Key", sPropKey)
                        .WriteElementString("Value", props.Value)
                        .WriteEndElement()
                    Next
                    For Each sDisplayedKey As String In chs.htblDisplayedDescriptions.Keys
                        .WriteElementString("Displayed", sDisplayedKey)
                    Next
                    .WriteEndElement()
                Next

                For Each sKey As String In state.htblVariableStates.Keys
                    Dim vars As clsGameState.clsVariableState = CType(state.htblVariableStates(sKey), clsGameState.clsVariableState)
                    .WriteStartElement("Variable")
                    .WriteElementString("Key", sKey)

                    Dim v As clsVariable = Adventure.htblVariables(sKey)
                    For i As Integer = 0 To vars.Value.Length - 1
                        If v.Type = clsVariable.VariableTypeEnum.Numeric Then
                            If vars.Value(i) <> "0" Then .WriteElementString("Value_" & i, vars.Value(i))
                        Else
                            If vars.Value(i) <> "" Then .WriteElementString("Value_" & i, vars.Value(i))
                        End If
                    Next
                    For Each sDisplayedKey As String In vars.htblDisplayedDescriptions.Keys
                        .WriteElementString("Displayed", sDisplayedKey)
                    Next
                    .WriteEndElement()
                Next

                For Each sKey As String In state.htblGroupStates.Keys
                    Dim grps As clsGameState.clsGroupState = CType(state.htblGroupStates(sKey), clsGameState.clsGroupState)
                    .WriteStartElement("Group")
                    .WriteElementString("Key", sKey)
                    For Each sMember As String In grps.lstMembers
                        .WriteElementString("Member", sMember)
                    Next
                    .WriteEndElement()
                Next

                .WriteElementString("Turns", Adventure.Turns.ToString)

                .WriteEndElement() ' Game
                .WriteEndDocument()
                .Flush()

                Dim outStream As New MemoryStream
                Dim zStream As New DeflaterOutputStream(outStream)

                Try
                    stmMemory.Position = 0
                    CopyStream(stmMemory, zStream)
                Finally
                    zStream.Close()
                    bData = outStream.ToArray
                    outStream.Close()
                End Try
            End With

            Dim stmFile As New IO.FileStream(sFilePath, FileMode.Create)
            Dim bw As New BinaryWriter(stmFile)

            bw.Write(bData)
            bw.Close()
            stmFile.Close()

            Return True

        Catch ex As Exception
            ErrMsg("Error saving game state", ex)
            Return False
        End Try

    End Function

    Friend Function LoadActions(ByVal nodContainerXML As XmlElement) As ActionArrayList

        Dim Actions As New ActionArrayList

        If nodContainerXML.Item("Actions") Is Nothing Then Return Actions

        For Each nod As XmlNode In nodContainerXML.Item("Actions").ChildNodes
            If TypeOf nod Is XmlElement Then
                Dim nodAct As XmlElement = CType(nod, XmlElement)
                Dim act As New clsAction
                Dim sAct As String
                Dim sType As String = Nothing
                sType = nodAct.Name

                If sType Is Nothing Then Return Actions

                sAct = nodAct.InnerText
                Dim sElements() As String = Split(sAct, " ")
                Select Case sType
                    Case "EndGame"
                        act.eItem = clsAction.ItemEnum.EndGame
                        act.eEndgame = EnumParseEndGame(sElements(0))

                    Case "MoveObject", "AddObjectToGroup", "RemoveObjectFromGroup"
                        Select Case sType
                            Case "MoveObject"
                                act.eItem = clsAction.ItemEnum.MoveObject
                            Case "AddObjectToGroup"
                                act.eItem = clsAction.ItemEnum.AddObjectToGroup
                            Case "RemoveObjectFromGroup"
                                act.eItem = clsAction.ItemEnum.RemoveObjectFromGroup
                        End Select

                        If dFileVersion <= 5.000016 Then ' Upgrade previous file format
                            act.sKey1 = sElements(0)
                            act.eMoveObjectWhat = clsAction.MoveObjectWhatEnum.Object
                            Select Case act.sKey1
                                Case "AllHeldObjects"
                                    act.eMoveObjectWhat = clsAction.MoveObjectWhatEnum.EverythingHeldBy
                                    act.sKey1 = THEPLAYER
                                Case "AllWornObjects"
                                    act.eMoveObjectWhat = clsAction.MoveObjectWhatEnum.EverythingWornBy
                                    act.sKey1 = THEPLAYER
                                Case Else
                                    ' Leave as is
                            End Select
                            act.eMoveObjectTo = EnumParseMoveObject(sElements(1))
                            act.sKey2 = sElements(2)
                        Else
                            act.eMoveObjectWhat = CType([Enum].Parse(GetType(clsAction.MoveObjectWhatEnum), sElements(0)), clsAction.MoveObjectWhatEnum)
                            act.sKey1 = sElements(1)
                            If sElements.Length > 4 Then
                                For iEl As Integer = 2 To sElements.Length - 3
                                    act.sPropertyValue &= sElements(iEl)
                                    If iEl < sElements.Length - 3 Then act.sPropertyValue &= " "
                                Next
                            End If
                            Select Case act.eItem
                                Case clsAction.ItemEnum.AddObjectToGroup
                                    act.eMoveObjectTo = clsAction.MoveObjectToEnum.ToGroup
                                Case clsAction.ItemEnum.RemoveObjectFromGroup
                                    act.eMoveObjectTo = clsAction.MoveObjectToEnum.FromGroup
                                Case clsAction.ItemEnum.MoveObject
                                    act.eMoveObjectTo = CType([Enum].Parse(GetType(clsAction.MoveObjectToEnum), sElements(sElements.Length - 2)), clsAction.MoveObjectToEnum)
                            End Select
                            act.sKey2 = sElements(sElements.Length - 1)
                        End If

                    Case "MoveCharacter", "AddCharacterToGroup", "RemoveCharacterFromGroup"
                        Select Case sType
                            Case "MoveCharacter"
                                act.eItem = clsAction.ItemEnum.MoveCharacter
                            Case "AddCharacterToGroup"
                                act.eItem = clsAction.ItemEnum.AddCharacterToGroup
                            Case "RemoveCharacterFromGroup"
                                act.eItem = clsAction.ItemEnum.RemoveCharacterFromGroup
                        End Select

                        If dFileVersion <= 5.000016 Then ' Upgrade previous file format
                            act.eItem = clsAction.ItemEnum.MoveCharacter
                            act.sKey1 = sElements(0)
                            act.eMoveCharacterTo = EnumParseMoveCharacter(sElements(1))
                            act.sKey2 = sElements(2)
                            If act.eMoveCharacterTo = clsAction.MoveCharacterToEnum.InDirection AndAlso IsNumeric(act.sKey2) Then
                                act.sKey2 = WriteEnum(CType(SafeInt(act.sKey2), DirectionsEnum))
                            End If
                        Else
                            act.eMoveCharacterWho = CType([Enum].Parse(GetType(clsAction.MoveCharacterWhoEnum), sElements(0)), clsAction.MoveCharacterWhoEnum)
                            act.sKey1 = sElements(1)
                            If sElements.Length > 4 Then
                                For iEl As Integer = 2 To sElements.Length - 3
                                    act.sPropertyValue &= sElements(iEl)
                                    If iEl < sElements.Length - 3 Then act.sPropertyValue &= " "
                                Next
                            End If
                            Select Case act.eItem
                                Case clsAction.ItemEnum.AddCharacterToGroup
                                    act.eMoveCharacterTo = clsAction.MoveCharacterToEnum.ToGroup
                                Case clsAction.ItemEnum.RemoveCharacterFromGroup
                                    act.eMoveCharacterTo = clsAction.MoveCharacterToEnum.FromGroup
                                Case clsAction.ItemEnum.MoveCharacter
                                    act.eMoveCharacterTo = CType([Enum].Parse(GetType(clsAction.MoveCharacterToEnum), sElements(sElements.Length - 2)), clsAction.MoveCharacterToEnum)
                            End Select
                            act.sKey2 = sElements(sElements.Length - 1)
                        End If

                    Case "AddLocationToGroup", "RemoveLocationFromGroup"
                        Select Case sType
                            Case "AddLocationToGroup"
                                act.eItem = clsAction.ItemEnum.AddLocationToGroup
                            Case "RemoveLocationFromGroup"
                                act.eItem = clsAction.ItemEnum.RemoveLocationFromGroup
                        End Select

                        act.eMoveLocationWhat = CType([Enum].Parse(GetType(clsAction.MoveLocationWhatEnum), sElements(0)), clsAction.MoveLocationWhatEnum)
                        act.sKey1 = sElements(1)
                        If sElements.Length > 4 Then
                            For iEl As Integer = 2 To sElements.Length - 3
                                act.sPropertyValue &= sElements(iEl)
                                If iEl < sElements.Length - 3 Then act.sPropertyValue &= " "
                            Next
                        End If
                        Select Case act.eItem
                            Case clsAction.ItemEnum.AddLocationToGroup
                                act.eMoveLocationTo = clsAction.MoveLocationToEnum.ToGroup
                            Case clsAction.ItemEnum.RemoveLocationFromGroup
                                act.eMoveLocationTo = clsAction.MoveLocationToEnum.FromGroup
                        End Select
                        act.sKey2 = sElements(sElements.Length - 1)

                    Case "SetProperty"
                        act.eItem = clsAction.ItemEnum.SetProperties
                        act.sKey1 = sElements(0)
                        act.sKey2 = sElements(1)
                        'act.StringValue = sElements(2)
                        Dim sValue As String = ""
                        For i As Integer = 2 To sElements.Length - 1
                            sValue &= sElements(i)
                            If i < sElements.Length - 1 Then sValue &= " "
                        Next
                        act.StringValue = sValue
                        act.sPropertyValue = sValue
                    Case "Score"
                    Case "SetTasks"
                        act.eItem = clsAction.ItemEnum.SetTasks
                        Dim iStartIndex As Integer = 0
                        Dim iEndIndex As Integer = 1
                        If sElements(0) = "FOR" Then
                            ' sElements(1) = %Loop%
                            ' sElements(2) = '='
                            act.IntValue = CInt(sElements(3))
                            ' sElements(4) = TO
                            act.sPropertyValue = sElements(5)
                            ' sElements(6) = :
                            iStartIndex = 7
                            iEndIndex = 4
                        End If
                        act.eSetTasks = EnumParseSetTask(sElements(iStartIndex))
                        act.sKey1 = sElements(iStartIndex + 1)
                        For iElement As Integer = iStartIndex + 2 To sElements.Length - iEndIndex
                            act.StringValue &= sElements(iElement)
                        Next
                        If act.StringValue IsNot Nothing Then
                            ' Simplify Runner so it only has to deal with multiple, or specific refs
                            act.StringValue = FixInitialRefs(act.StringValue)
                            If act.StringValue.StartsWith("(") Then act.StringValue = sRight(act.StringValue, act.StringValue.Length - 1)
                            If act.StringValue.EndsWith(")") Then act.StringValue = sLeft(act.StringValue, act.StringValue.Length - 1)
                        End If

                    Case "SetVariable", "IncVariable", "DecVariable", "ExecuteTask"
                        Select Case sType
                            Case "SetVariable"
                                act.eItem = clsAction.ItemEnum.SetVariable
                            Case "IncVariable"
                                act.eItem = clsAction.ItemEnum.IncreaseVariable
                            Case "DecVariable"
                                act.eItem = clsAction.ItemEnum.DecreaseVariable
                        End Select

                        If sElements(0) = "FOR" Then
                            act.eVariables = clsAction.VariablesEnum.Loop
                            ' sElements(1) = %Loop%
                            ' sElements(2) = '='
                            act.IntValue = CInt(sElements(3))
                            ' sElements(4) = TO
                            act.sKey2 = sElements(5)
                            ' sElements(6) = :
                            ' sElements(7) = SET
                            act.sKey1 = sElements(8).Split("["c)(0)
                            ' sElements(9) = '='
                            For iElement As Integer = 10 To sElements.Length - 4
                                act.StringValue &= sElements(iElement)
                                If iElement < sElements.Length - 4 Then act.StringValue &= " "
                            Next
                        Else
                            act.eVariables = clsAction.VariablesEnum.Assignment
                            If sInstr(sElements(0), "[") > 0 Then
                                act.sKey1 = sElements(0).Split("["c)(0)
                                act.sKey2 = sElements(0).Split("["c)(1).Replace("]", "")
                            Else
                                act.sKey1 = sElements(0)
                            End If
                            ' sElements(1) = '='
                            'act.StringValue = sElements(2)
                            For iElement As Integer = 2 To sElements.Length - 1
                                act.StringValue &= sElements(iElement)
                                If iElement < sElements.Length - 1 Then act.StringValue &= " "
                            Next
                            If dFileVersion > 5.0000321 Then
                                If act.StringValue.StartsWith("""") AndAlso act.StringValue.EndsWith("""") Then
                                    act.StringValue = act.StringValue.Substring(1, act.StringValue.Length - 2)
                                End If
                            End If
                        End If

                    Case "Time"
                        act.eItem = clsAction.ItemEnum.Time
                        For i As Integer = 1 To sElements.Length - 2
                            If i > 1 Then act.StringValue &= " "
                            act.StringValue &= sElements(i)
                        Next
                        act.StringValue = act.StringValue.Substring(1, act.StringValue.Length - 2)

                    Case "Conversation"
                        act.eItem = clsAction.ItemEnum.Conversation
                        Select Case sElements(0).ToUpper
                            Case "GREET", "FAREWELL"
                                If sElements(0).ToUpper = "GREET" Then
                                    act.eConversation = clsAction.ConversationEnum.Greet
                                Else
                                    act.eConversation = clsAction.ConversationEnum.Farewell
                                End If
                                act.sKey1 = sElements(1)
                                If sElements.Length > 2 Then
                                    ' sElements(2) = "with"
                                    For iElement As Integer = 3 To sElements.Length - 1
                                        act.StringValue &= sElements(iElement)
                                        If iElement < sElements.Length - 1 Then act.StringValue &= " "
                                    Next
                                    If act.StringValue.StartsWith("'") Then act.StringValue = act.StringValue.Substring(1)
                                    If act.StringValue.EndsWith("'") Then act.StringValue = sLeft(act.StringValue, act.StringValue.Length - 1)
                                End If

                            Case "ASK", "TELL"
                                If sElements(0).ToUpper = "ASK" Then
                                    act.eConversation = clsAction.ConversationEnum.Ask
                                Else
                                    act.eConversation = clsAction.ConversationEnum.Tell
                                End If
                                act.sKey1 = sElements(1)
                                ' sElements(2) = "About"
                                For iElement As Integer = 3 To sElements.Length - 1
                                    act.StringValue &= sElements(iElement)
                                    If iElement < sElements.Length - 1 Then act.StringValue &= " "
                                Next
                                If act.StringValue.StartsWith("'") Then act.StringValue = act.StringValue.Substring(1)
                                If act.StringValue.EndsWith("'") Then act.StringValue = sLeft(act.StringValue, act.StringValue.Length - 1)

                            Case "SAY"
                                act.eConversation = clsAction.ConversationEnum.Command
                                ' sElements(0) = "Say"
                                For iElement As Integer = 1 To sElements.Length - 3
                                    act.StringValue &= sElements(iElement)
                                    If iElement < sElements.Length - 3 Then act.StringValue &= " "
                                Next
                                If act.StringValue.StartsWith("'") Then act.StringValue = act.StringValue.Substring(1)
                                If act.StringValue.EndsWith("'") Then act.StringValue = sLeft(act.StringValue, act.StringValue.Length - 1)
                                ' sElements(len - 2) = "to"
                                act.sKey1 = sElements(sElements.Length - 1)

                            Case "ENTERWITH", "LEAVEWITH"
                                If sElements(0).ToUpper = "ENTERWITH" Then
                                    act.eConversation = clsAction.ConversationEnum.EnterConversation
                                Else
                                    act.eConversation = clsAction.ConversationEnum.LeaveConversation
                                End If
                                act.sKey1 = sElements(1)
                        End Select
                End Select

                Actions.Add(act)

            End If
        Next

        Return Actions

    End Function


    Friend Function LoadRestrictions(ByVal nodContainerXML As XmlElement) As RestrictionArrayList

        Dim Restrictions As New RestrictionArrayList

        If nodContainerXML.SelectNodes("Restrictions").Count = 0 Then Return Restrictions

        Dim nodRestrictions As XmlElement = CType(nodContainerXML.SelectNodes("Restrictions")(0), XmlElement)

        For Each nodRest As XmlElement In nodRestrictions.SelectNodes("Restriction")
            Dim rest As New clsRestriction
            Dim sRest As String
            Dim sType As String = Nothing
            If Not nodRest.Item("Location") Is Nothing Then sType = "Location"
            If Not nodRest.Item("Object") Is Nothing Then sType = "Object"
            If Not nodRest.Item("Task") Is Nothing Then sType = "Task"
            If Not nodRest.Item("Character") Is Nothing Then sType = "Character"
            If Not nodRest.Item("Variable") Is Nothing Then sType = "Variable"
            If Not nodRest.Item("Property") Is Nothing Then sType = "Property"
            If nodRest.Item("Direction") IsNot Nothing Then sType = "Direction"
            If nodRest.Item("Expression") IsNot Nothing Then sType = "Expression"
            If nodRest.Item("Item") IsNot Nothing Then sType = "Item"

            If nodRest.Item(sType) Is Nothing Then Exit For

            sRest = nodRest.Item(sType).InnerText
            Dim sElements() As String = Split(sRest, " ")
            Select Case sType
                Case "Location"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Location
                    rest.sKey1 = sElements(0)
                    rest.eMust = EnumParseMust(sElements(1))

                    rest.eLocation = EnumParseLocation(sElements(2))
                    rest.sKey2 = sElements(3)
                Case "Object"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Object
                    rest.sKey1 = sElements(0)
                    rest.eMust = EnumParseMust(sElements(1))
                    rest.eObject = EnumParseObject(sElements(2))
                    rest.sKey2 = sElements(3)
                Case "Task"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Task
                    rest.sKey1 = sElements(0)
                    rest.eMust = EnumParseMust(sElements(1))
                    rest.eTask = clsRestriction.TaskEnum.Complete
                Case "Character"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Character
                    rest.sKey1 = sElements(0)
                    rest.eMust = EnumParseMust(sElements(1))
                    rest.eCharacter = EnumParseCharacter(sElements(2))
                    rest.sKey2 = sElements(3)
                Case "Item"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Item
                    rest.sKey1 = sElements(0)
                    rest.eMust = EnumParseMust(sElements(1))
                    rest.eItem = EnumParseItem(sElements(2))
                    rest.sKey2 = sElements(3)
                Case "Variable"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Variable
                    rest.sKey1 = sElements(0)
                    If rest.sKey1.Contains("[") AndAlso rest.sKey1.Contains("]") Then
                        rest.sKey2 = rest.sKey1.Substring(rest.sKey1.IndexOf("[") + 1, rest.sKey1.LastIndexOf("]") - rest.sKey1.IndexOf("[") - 1)
                        rest.sKey1 = rest.sKey1.Substring(0, rest.sKey1.IndexOf("["))
                    End If
                    rest.eMust = EnumParseMust(sElements(1))
                    rest.eVariable = CType([Enum].Parse(GetType(clsRestriction.VariableEnum), sElements(2).Substring(2)), clsRestriction.VariableEnum)

                    Dim sValue As String = ""
                    For i As Integer = 3 To sElements.Length - 1
                        sValue &= sElements(i)
                        If i < sElements.Length - 1 Then sValue &= " "
                    Next
                    If sElements.Length = 4 AndAlso IsNumeric(sElements(3)) Then
                        rest.IntValue = CInt(sElements(3)) ' Integer value
                        rest.StringValue = rest.IntValue.ToString
                    Else
                        If sValue.StartsWith("""") AndAlso sValue.EndsWith("""") Then
                            rest.StringValue = sValue.Substring(1, sValue.Length - 2) ' String constant
                        ElseIf sValue.StartsWith("'") AndAlso sValue.EndsWith("'") Then
                            rest.StringValue = sValue.Substring(1, sValue.Length - 2) ' Expression
                        Else
                            rest.StringValue = sElements(3)
                            rest.IntValue = Integer.MinValue ' A key to a variable
                        End If
                    End If

                Case "Property"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Property
                    rest.sKey1 = sElements(0)
                    rest.sKey2 = sElements(1)
                    rest.eMust = EnumParseMust(sElements(2))
                    Dim iStartExpression As Integer = 3
                    rest.IntValue = -1
                    For Each eEquals As clsRestriction.VariableEnum In [Enum].GetValues(GetType(clsRestriction.VariableEnum))
                        If sElements(3) = eEquals.ToString Then rest.IntValue = CInt(eEquals)
                    Next
                    If rest.IntValue > -1 Then iStartExpression = 4 Else rest.IntValue = CInt(clsRestriction.VariableEnum.EqualTo)
                    Dim sValue As String = ""
                    For i As Integer = iStartExpression To sElements.Length - 1
                        sValue &= sElements(i)
                        If i < sElements.Length - 1 Then sValue &= " "
                    Next
                    rest.StringValue = sValue

                Case "Direction"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Direction
                    rest.eMust = EnumParseMust(sElements(0))
                    rest.sKey1 = sElements(1)
                    rest.sKey1 = sRight(rest.sKey1, rest.sKey1.Length - 2) ' Trim off the Be

                Case "Expression"
                    rest.eType = clsRestriction.RestrictionTypeEnum.Expression
                    rest.eMust = clsRestriction.MustEnum.Must
                    Dim sValue As String = ""
                    For i As Integer = 0 To sElements.Length - 1
                        sValue &= sElements(i)
                        If i < sElements.Length - 1 Then sValue &= " "
                    Next
                    rest.StringValue = sValue

            End Select
            rest.oMessage = LoadDescription(nodRest, "Message")
            Restrictions.Add(rest)

        Next
        Restrictions.BracketSequence = nodRestrictions.SelectNodes("BracketSequence")(0).InnerText ' GetElementsByTagName("BracketSequence")(0).InnerText

        If Not bAskedAboutBrackets AndAlso dFileVersion < 5.000026 AndAlso Restrictions.BracketSequence.Contains("#A#O#") Then
            bCorrectBracketSequences = Glue.AskYesNoQuestion("There was a logic correction in version 5.0.26 which means OR restrictions after AND restrictions were not evaluated.  Would you like to auto-correct these tasks?" & vbCrLf & vbCrLf & "You may not wish to do so if you have already used brackets around any OR restrictions following AND restrictions.", "Adventure Upgrade")
            bAskedAboutBrackets = True
        End If
        If bCorrectBracketSequences Then Restrictions.BracketSequence = CorrectBracketSequence(Restrictions.BracketSequence)

        Restrictions.BracketSequence = Restrictions.BracketSequence.Replace("[", "((").Replace("]", "))")
        Return Restrictions

    End Function


    Private Function FixInitialRefs(ByVal sCommand As String) As String
        If sCommand Is Nothing Then Return ""
        Return sCommand.Replace("%object%", "%object1%").Replace("%character%", "%character1%").Replace("%location%", "%location1%").Replace("%number%", "%number1%").Replace("%text%", "%text1%").Replace("%item%", "%item1%").Replace("%direction%", "%direction1%")
    End Function


    Friend Sub SaveRestrictions(ByRef xmlWriter As XmlTextWriter, ByVal Restrictions As RestrictionArrayList)

        If Restrictions.Count = 0 Then Exit Sub

        With xmlWriter
            .WriteStartElement("Restrictions")
            For Each rest As clsRestriction In Restrictions
                .WriteStartElement("Restriction")
                Select Case rest.eType
                    Case clsRestriction.RestrictionTypeEnum.Location
                        .WriteElementString("Location", rest.sKey1 & " " & WriteEnum(rest.eMust) & " " & WriteEnum(rest.eLocation) & " " & rest.sKey2)
                    Case clsRestriction.RestrictionTypeEnum.Object
                        .WriteElementString("Object", rest.sKey1 & " " & WriteEnum(rest.eMust) & " " & WriteEnum(rest.eObject) & " " & rest.sKey2)
                    Case clsRestriction.RestrictionTypeEnum.Task
                        .WriteElementString("Task", rest.sKey1 & " " & WriteEnum(rest.eMust) & " BeComplete")
                    Case clsRestriction.RestrictionTypeEnum.Character
                        .WriteElementString("Character", rest.sKey1 & " " & WriteEnum(rest.eMust) & " " & WriteEnum(rest.eCharacter) & " " & rest.sKey2)
                    Case clsRestriction.RestrictionTypeEnum.Item
                        .WriteElementString("Item", rest.sKey1 & " " & WriteEnum(rest.eMust) & " " & WriteEnum(rest.eItem) & " " & rest.sKey2)
                    Case clsRestriction.RestrictionTypeEnum.Variable
                        Dim sValue As String
                        Dim eVarType As clsVariable.VariableTypeEnum
                        If rest.sKey1.StartsWith("ReferencedNumber") Then
                            eVarType = clsVariable.VariableTypeEnum.Numeric
                        ElseIf rest.sKey1.StartsWith("ReferencedText") Then
                            eVarType = clsVariable.VariableTypeEnum.Text
                        Else
                            Dim var As clsVariable = Adventure.htblVariables(rest.sKey1)
                            If var IsNot Nothing Then eVarType = var.Type
                        End If

                        If rest.IntValue = Integer.MinValue Then
                            sValue = rest.StringValue ' Key to variable
                        Else
                            If eVarType = clsVariable.VariableTypeEnum.Numeric Then
                                If rest.StringValue <> "" AndAlso rest.StringValue <> rest.IntValue.ToString Then
                                    sValue = "'" & rest.StringValue & "'" ' Expression
                                Else
                                    sValue = rest.IntValue.ToString ' Integer
                                End If
                            Else
                                sValue = """" & rest.StringValue & """" ' String constant
                            End If
                        End If
                        Dim sVar As String = rest.sKey1
                        If rest.sKey2 <> "" Then
                            If Adventure.htblVariables.ContainsKey(rest.sKey2) Then
                                sVar &= "[%" & Adventure.htblVariables(rest.sKey2).Name & "%]"
                            Else
                                sVar &= "[" & rest.sKey2 & "]"
                            End If
                        End If
                        .WriteElementString("Variable", sVar & " " & WriteEnum(rest.eMust) & " Be" & rest.eVariable.ToString & " " & sValue)

                    Case clsRestriction.RestrictionTypeEnum.Property
                        Dim sEquals As String = ""
                        If rest.IntValue > -1 Then sEquals = CType(rest.IntValue, clsRestriction.VariableEnum).ToString & " "
                        .WriteElementString("Property", rest.sKey1 & " " & rest.sKey2 & " " & WriteEnum(rest.eMust) & " " & sEquals & rest.StringValue)
                    Case clsRestriction.RestrictionTypeEnum.Direction
                        .WriteElementString("Direction", WriteEnum(rest.eMust) & " Be" & rest.sKey1)
                    Case clsRestriction.RestrictionTypeEnum.Expression
                        .WriteElementString("Expression", rest.StringValue)
                End Select
                If rest.oMessage.Count > 0 Then SaveDescription(xmlWriter, rest.oMessage, "Message") ' .WriteElementString("Message", rest.sMessage)
                .WriteEndElement() ' Restriction
            Next
            .WriteElementString("BracketSequence", Restrictions.BracketSequence)
            .WriteEndElement()
        End With


    End Sub


    Friend Sub SaveActions(ByRef xmlWriter As XmlTextWriter, ByVal Actions As ActionArrayList)

        If Actions.Count = 0 Then Exit Sub

        With xmlWriter
            .WriteStartElement("Actions")
            For Each act As clsAction In Actions
                Select Case act.eItem
                    Case clsAction.ItemEnum.EndGame
                        .WriteElementString("EndGame", WriteEnum(act.eEndgame))
                    Case clsAction.ItemEnum.MoveCharacter, clsAction.ItemEnum.AddCharacterToGroup, clsAction.ItemEnum.RemoveCharacterFromGroup
                        .WriteElementString(act.eItem.ToString, act.eMoveCharacterWho.ToString & " " & act.sKey1 & " " & IIf(act.sPropertyValue <> "", act.sPropertyValue & " ", "").ToString & WriteEnum(act.eMoveCharacterTo) & " " & act.sKey2)
                    Case clsAction.ItemEnum.MoveObject, clsAction.ItemEnum.AddObjectToGroup, clsAction.ItemEnum.RemoveObjectFromGroup
                        .WriteElementString(act.eItem.ToString, act.eMoveObjectWhat.ToString & " " & act.sKey1 & " " & IIf(act.sPropertyValue <> "", act.sPropertyValue & " ", "").ToString & WriteEnum(act.eMoveObjectTo) & " " & act.sKey2)
                    Case clsAction.ItemEnum.AddLocationToGroup, clsAction.ItemEnum.RemoveLocationFromGroup
                        .WriteElementString(act.eItem.ToString, act.eMoveLocationWhat.ToString & " " & act.sKey1 & " " & IIf(act.sPropertyValue <> "", act.sPropertyValue & " ", "").ToString & WriteEnum(act.eMoveLocationTo) & " " & act.sKey2)
                    Case clsAction.ItemEnum.SetProperties
                        .WriteElementString("SetProperty", act.sKey1 & " " & act.sKey2 & " " & act.sPropertyValue)
                    Case clsAction.ItemEnum.SetTasks
                        Dim sAction As String = ""
                        If act.sPropertyValue <> "" Then sAction = "FOR Loop = " & act.IntValue & " TO " & act.sPropertyValue & " : "
                        Dim sParams As String = ""
                        If act.StringValue <> "" Then sParams = " (" & act.StringValue & ")"
                        sAction &= WriteEnum(act.eSetTasks) & " " & act.sKey1 & sParams
                        If act.sPropertyValue <> "" Then sAction &= " : NEXT Loop"
                        .WriteElementString("SetTasks", sAction)
                    Case clsAction.ItemEnum.SetVariable, clsAction.ItemEnum.IncreaseVariable, clsAction.ItemEnum.DecreaseVariable
                        Dim var As clsVariable = Adventure.htblVariables(act.sKey1)
                        Dim sAction As String
                        If act.eVariables = clsAction.VariablesEnum.Assignment Then
                            sAction = act.sKey1
                            If Not act.sKey2 Is Nothing Then
                                sAction &= "[" & act.sKey2 & "]"
                            End If
                            sAction &= " = """ & act.StringValue & """"
                        Else
                            sAction = "FOR Loop = " & act.IntValue & " TO " & act.sKey2 & " : SET " & act.sKey1 & "[Loop] = " & act.StringValue & " : NEXT Loop"
                        End If
                        Select Case act.eItem
                            Case clsAction.ItemEnum.SetVariable
                                .WriteElementString("SetVariable", sAction)
                            Case clsAction.ItemEnum.IncreaseVariable
                                .WriteElementString("IncVariable", sAction)
                            Case clsAction.ItemEnum.DecreaseVariable
                                .WriteElementString("DecVariable", sAction)
                        End Select
                    Case clsAction.ItemEnum.Conversation
                        Select Case act.eConversation
                            Case clsAction.ConversationEnum.Greet
                                .WriteElementString("Conversation", "Greet " & act.sKey1 & IIf(act.StringValue <> "", " With '" & act.StringValue & "'", "").ToString)
                            Case clsAction.ConversationEnum.Ask
                                .WriteElementString("Conversation", "Ask " & act.sKey1 & " About '" & act.StringValue & "'")
                            Case clsAction.ConversationEnum.Tell
                                .WriteElementString("Conversation", "Tell " & act.sKey1 & " About '" & act.StringValue & "'")
                            Case clsAction.ConversationEnum.Command
                                .WriteElementString("Conversation", "Say '" & act.StringValue & "' To " & act.sKey1)
                            Case clsAction.ConversationEnum.Farewell
                                .WriteElementString("Conversation", "Farewell " & act.sKey1 & IIf(act.StringValue <> "", " With '" & act.StringValue & "'", "").ToString)
                            Case clsAction.ConversationEnum.EnterConversation
                                .WriteElementString("Conversation", "EnterWith " & act.sKey1)
                            Case clsAction.ConversationEnum.LeaveConversation
                                .WriteElementString("Conversation", "LeaveWith " & act.sKey1)
                        End Select
                    Case clsAction.ItemEnum.Time
                        .WriteElementString("Time", "Skip """ & act.StringValue & """ turns")
                End Select
            Next
            .WriteEndElement()
        End With


    End Sub


    Private Function IsEqual(ByRef mb1 As Byte(), ByRef mb2 As Byte()) As Boolean

        If (mb1.Length <> mb2.Length) Then ' make sure arrays same length
            Return False
        Else
            For i As Integer = 0 To mb1.Length - 1 ' run array length looking for miscompare
                If mb1(i) <> mb2(i) Then Return False
            Next
        End If

        Return True
    End Function

    Private Function IsHex(ByVal sHex As String) As Boolean
        For i As Integer = 0 To sHex.Length - 1
            If "0123456789ABCDEF".IndexOf(sHex(i)) = -1 Then Return False
        Next
        Return True
    End Function


    ' "Tweak" any v5 library tasks that are different from v4
    Private Sub TweakTasksForv4()

        If Adventure.htblTasks.ContainsKey("GiveObjectToChar") Then
            With Adventure.htblTasks("GiveObjectToChar")
                .CompletionMessage = New Description("%CharacterName[%character%, subject]% doesn't seem interested in %objects%.Name.")
                .arlActions.Clear()
            End With
        End If

        If Adventure.htblTasks.ContainsKey("Look") Then
            With Adventure.htblTasks("Look")
                .arlCommands(0) = "[look/l]{ room}"
                .arlCommands.Add("[x/examine] room")
            End With
        End If

        SearchAndReplace("Sorry, I'm not sure which object or character you are trying to examine.", "You see no such thing.", True)

    End Sub


    ' This loads from file into our data structure
    ' Assumes file exists
    '
    Public Enum FileTypeEnum
        TextAdventure_TAF
        XMLModule_AMF
        v4Module_AMF
        GameState_TAS
        Blorb
        Exe
    End Enum

    Friend Function LoadFile(ByVal sFilename As String, ByVal eFileType As FileTypeEnum, ByVal eLoadWhat As LoadWhatEnum, ByVal bLibrary As Boolean, Optional ByVal dtAdvDate As Date = #1/1/1900#, Optional ByRef state As clsGameState = Nothing, Optional ByVal lOffset As Long = 0, Optional ByVal bSilentError As Boolean = False) As Boolean

        Dim stmOriginalFile As IO.FileStream = Nothing

        Try
            If Not IO.File.Exists(sFilename) Then
                ErrMsg("File '" & sFilename & "' not found.")
                RemoveFileFromList(sFilename)
                Return False
            End If

            stmOriginalFile = New IO.FileStream(sFilename, IO.FileMode.Open, FileAccess.Read)

            If eFileType = FileTypeEnum.TextAdventure_TAF Then

                stmOriginalFile.Seek(FileLen(sFilename) - 14, SeekOrigin.Begin)
                br = New BinaryReader(stmOriginalFile)
                Dim bPass As Byte() = Dencode(br.ReadBytes(12), FileLen(sFilename) - 13)
                Dim sPassString As String = System.Text.Encoding.UTF8.GetString(bPass)
                Adventure.Password = (sLeft(sPassString, 4) & sRight(sPassString, 4)).Trim

            End If

            If lOffset > 0 Then
                stmOriginalFile.Seek(lOffset, SeekOrigin.Begin)
                lOffset += 7 ' for the footer
            Else
                stmOriginalFile.Seek(0, SeekOrigin.Begin)
            End If

            br = New BinaryReader(stmOriginalFile)

            iLoading += 1

            Select Case eFileType
                Case FileTypeEnum.Exe
                    If clsBlorb.ExecResource IsNot Nothing AndAlso clsBlorb.ExecResource.Length > 0 Then
                        If Not Load500(Decompress(clsBlorb.ExecResource, dVersion >= 5.00002 AndAlso clsBlorb.bObfuscated, 16, clsBlorb.ExecResource.Length - 30), False) Then Return False
                        clsBlorb.bObfuscated = False
                        If clsBlorb.MetaData IsNot Nothing Then Adventure.BabelTreatyInfo.FromString(clsBlorb.MetaData.OuterXml)
                        Adventure.FullPath = SafeString(Application.ExecutablePath)
                    Else
                        Return False
                    End If

                Case FileTypeEnum.Blorb
                    Blorb = New clsBlorb
                    If Blorb.LoadBlorb(stmOriginalFile, sFilename) Then

                        Dim bVersion As Byte() = New Byte(11) {}
                        Array.Copy(clsBlorb.ExecResource, bVersion, 12)

                        Dim sVersion As String
                        If IsEqual(bVersion, New Byte() {60, 66, 63, 201, 106, 135, 194, 207, 146, 69, 62, 97}) Then
                            sVersion = "Version 5.00"
                        ElseIf IsEqual(bVersion, New Byte() {60, 66, 63, 201, 106, 135, 194, 207, 147, 69, 62, 97}) Then
                            sVersion = "Version 4.00"
                        ElseIf IsEqual(bVersion, New Byte() {60, 66, 63, 201, 106, 135, 194, 207, 148, 69, 55, 97}) Then
                            sVersion = "Version 3.90"
                        Else
                            sVersion = System.Text.Encoding.UTF8.GetString(Dencode(bVersion, 1))
                        End If

                        If Left(sVersion, 8) <> "Version " Then
                            ErrMsg("Not an ADRIFT Blorb file")
                            Return False
                        End If

                        UserSession.ShowUserSplash()

                        With Adventure
                            .dVersion = Double.Parse(sVersion.Replace("Version ", ""), Globalization.CultureInfo.InvariantCulture.NumberFormat) 'CDbl(sVersion.Replace("Version ", ""))
                            .Filename = Path.GetFileName(sFilename)
                            .FullPath = sFilename
                        End With

                        Dim bDeObfuscate As Boolean = clsBlorb.MetaData Is Nothing OrElse clsBlorb.MetaData.OuterXml.Contains("compilerversion") ' Nasty, but works
                        ' Was this a pre-obfuscated size blorb?
                        If clsBlorb.ExecResource.Length > 16 AndAlso clsBlorb.ExecResource(12) = 48 AndAlso clsBlorb.ExecResource(13) = 48 AndAlso clsBlorb.ExecResource(14) = 48 AndAlso clsBlorb.ExecResource(15) = 48 Then
                            If Not Load500(Decompress(clsBlorb.ExecResource, bDeObfuscate, 16, clsBlorb.ExecResource.Length - 30), False) Then Return False
                        Else
                            If Not Load500(Decompress(clsBlorb.ExecResource, bDeObfuscate, 12, clsBlorb.ExecResource.Length - 26), False) Then Return False
                        End If

                        If clsBlorb.MetaData IsNot Nothing Then Adventure.BabelTreatyInfo.FromString(clsBlorb.MetaData.OuterXml)
                    Else
                        Return False
                    End If

                Case FileTypeEnum.TextAdventure_TAF
                    Dim bVersion As Byte() = br.ReadBytes(12)
                    Dim sVersion As String
                    If IsEqual(bVersion, New Byte() {60, 66, 63, 201, 106, 135, 194, 207, 146, 69, 62, 97}) Then
                        sVersion = "Version 5.00"
                    ElseIf IsEqual(bVersion, New Byte() {60, 66, 63, 201, 106, 135, 194, 207, 147, 69, 62, 97}) Then
                        sVersion = "Version 4.00"
                    ElseIf IsEqual(bVersion, New Byte() {60, 66, 63, 201, 106, 135, 194, 207, 148, 69, 55, 97}) Then
                        sVersion = "Version 3.90"
                    Else
                        sVersion = System.Text.Encoding.UTF8.GetString(Dencode(bVersion, 1))
                    End If


                    If Left(sVersion, 8) <> "Version " Then
                        ErrMsg("Not an ADRIFT Text Adventure file")
                        Return False
                    End If

                    With Adventure
                        .dVersion = Double.Parse(sVersion.Replace("Version ", ""), Globalization.CultureInfo.InvariantCulture.NumberFormat)
                        .Filename = Path.GetFileName(sFilename)
                        .FullPath = sFilename
                    End With

                    Debug.WriteLine("Start Load: " & Now)
                    Select Case sVersion
                        Case "Version 3.90", "Version 4.00"
                            Dim br2 As BinaryReader = br
                            'LoadDefaults() ' If mandatory properties like StaticOrDynamic don't exist, create them                                                        
                            LoadLibraries(LoadWhatEnum.Properies)
                            br = br2
                            iStartPriority = 0
                            If LoadOlder(CDbl(sVersion.Substring(8))) Then
                                iStartPriority = 50000
                                LoadLibraries(LoadWhatEnum.AllExceptProperties, "standardlibrary")
                                'CreateMandatoryProperties()
                                TweakTasksForv4()
                                br = br2
                                Set400SpecificTasks()
                                Adventure.dVersion = CDbl(sVersion.Substring(8))
                                If Adventure.dVersion = 4 Then Adventure.dVersion = 4.000052
                            Else
                                Return False
                            End If
                        Case "Version 5.00"
                            Dim sSize As String = System.Text.Encoding.UTF8.GetString(br.ReadBytes(4))
                            Dim sCheck As String = System.Text.Encoding.UTF8.GetString(br.ReadBytes(8))
                            Dim iBabelLength As Integer = 0
                            Dim sBabel As String = Nothing
                            Dim bObfuscate As Boolean = True
                            If sSize = "0000" OrElse sCheck = "<ifindex" Then
                                stmOriginalFile.Seek(16, 0) ' Set to just after the size chunk
                                ' 5.0.20 format onwards
                                iBabelLength = CInt("&H" & sSize)
                                If iBabelLength > 0 Then
                                    Dim bBabel() As Byte = br.ReadBytes(iBabelLength)
                                    sBabel = System.Text.Encoding.UTF8.GetString(bBabel)
                                End If
                                iBabelLength += 4 ' For size header
                            Else
                                ' Pre 5.0.20 
                                ' THIS COULD BE AN EXTRACTED TAF, THEREFORE NO METADATA!!!
                                ' Ok, we have no uncompressed Babel info at the start.  Start over...
                                stmOriginalFile.Seek(0, SeekOrigin.Begin)
                                br = New BinaryReader(stmOriginalFile)
                                br.ReadBytes(12)
                                bObfuscate = False
                            End If
                            Dim stmFile As MemoryStream = FileToMemoryStream(True, CInt(FileLen(sFilename) - 26 - lOffset - iBabelLength), bObfuscate)
                            If stmFile Is Nothing Then Return False
                            If Not Load500(stmFile, False, False, eLoadWhat, dtAdvDate) Then Return False ' - 12)))
                            If sBabel <> "" Then
                                Adventure.BabelTreatyInfo.FromString(sBabel)
                                Dim sTemp As String = Adventure.CoverFilename
                                Adventure.CoverFilename = Nothing
                                Adventure.CoverFilename = sTemp ' Just to re-set the image in the Babel structure
                            End If

                        Case Else
                            ErrMsg("ADRIFT " & sVersion & " Adventures are not currently supported in ADRIFT v" & dVersion.ToString("0.0"))
                            Return False
                    End Select
                    Debug.WriteLine("End Load: " & Now)

                Case FileTypeEnum.v4Module_AMF
                    TODO("Version 4.0 Modules")

                Case FileTypeEnum.XMLModule_AMF
                    If Not Load500(FileToMemoryStream(False, CInt(FileLen(sFilename)), False), bLibrary, True, eLoadWhat, dtAdvDate, sFilename) Then Return False

                Case FileTypeEnum.GameState_TAS
                    state = LoadState(FileToMemoryStream(True, CInt(FileLen(sFilename)), False))

            End Select

            If br IsNot Nothing Then br.Close()
            br = Nothing
            stmOriginalFile.Close()
            stmOriginalFile = Nothing

            If Adventure.NotUnderstood = "" Then Adventure.NotUnderstood = "Sorry, I didn't understand that command."
            If Adventure.Player IsNot Nothing AndAlso (Adventure.Player.Location.ExistWhere = clsCharacterLocation.ExistsWhereEnum.Hidden OrElse Adventure.Player.Location.Key = "") Then
                If Adventure.htblLocations.Count > 0 Then
                    Dim locFirst As clsCharacterLocation = Adventure.Player.Location
                    locFirst.ExistWhere = clsCharacterLocation.ExistsWhereEnum.AtLocation
                    For Each sKey As String In Adventure.htblLocations.Keys
                        locFirst.Key = sKey
                        Exit For
                    Next
                    Adventure.Player.Move(locFirst)
                End If
            End If

            iLoading -= 1

            Return True
        Catch ex As Exception
            If Not br Is Nothing Then br.Close()
            If Not stmOriginalFile Is Nothing Then stmOriginalFile.Close()
            ErrMsg("Error loading " & sFilename, ex)
            Return False
        Finally
        End Try

    End Function


    Private Function LoadState(ByVal stmMemory As MemoryStream) As clsGameState

        Try
            Dim NewState As New clsGameState
            Dim xmlDoc As New XmlDocument
            xmlDoc.Load(stmMemory)

            With xmlDoc.Item("Game")
                For Each nodLoc As XmlElement In xmlDoc.SelectNodes("/Game/Location")
                    With nodLoc
                        Dim locs As New clsGameState.clsLocationState
                        Dim sKey As String = .Item("Key").InnerText
                        For Each nodProp As XmlElement In nodLoc.SelectNodes("Property")
                            Dim props As New clsGameState.clsLocationState.clsStateProperty
                            Dim sPropKey As String = nodProp.Item("Key").InnerText
                            props.Value = nodProp.Item("Value").InnerText
                            locs.htblProperties.Add(sPropKey, props)
                        Next
                        For Each nodDisplayed As XmlElement In nodLoc.SelectNodes("Displayed")
                            locs.htblDisplayedDescriptions.Add(nodDisplayed.InnerText, True)
                        Next
                        NewState.htblLocationStates.Add(sKey, locs)
                    End With
                Next

                For Each nodOb As XmlElement In xmlDoc.SelectNodes("/Game/Object")
                    With nodOb
                        Dim obs As New clsGameState.clsObjectState
                        obs.Location = New clsObjectLocation
                        Dim sKey As String = .Item("Key").InnerText
                        If .Item("DynamicExistWhere") IsNot Nothing Then
                            obs.Location.DynamicExistWhere = CType([Enum].Parse(GetType(clsObjectLocation.DynamicExistsWhereEnum), .Item("DynamicExistWhere").InnerText), clsObjectLocation.DynamicExistsWhereEnum)
                        Else
                            obs.Location.DynamicExistWhere = clsObjectLocation.DynamicExistsWhereEnum.Hidden
                        End If
                        If .Item("StaticExistWhere") IsNot Nothing Then
                            obs.Location.StaticExistWhere = CType([Enum].Parse(GetType(clsObjectLocation.StaticExistsWhereEnum), .Item("StaticExistWhere").InnerText), clsObjectLocation.StaticExistsWhereEnum)
                        Else
                            obs.Location.StaticExistWhere = clsObjectLocation.StaticExistsWhereEnum.NoRooms
                        End If
                        If .Item("LocationKey") IsNot Nothing Then obs.Location.Key = .Item("LocationKey").InnerText
                        For Each nodProp As XmlElement In nodOb.SelectNodes("Property")
                            Dim props As New clsGameState.clsObjectState.clsStateProperty
                            Dim sPropKey As String = nodProp.Item("Key").InnerText
                            props.Value = nodProp.Item("Value").InnerText
                            obs.htblProperties.Add(sPropKey, props)
                        Next
                        For Each nodDisplayed As XmlElement In nodOb.SelectNodes("Displayed")
                            obs.htblDisplayedDescriptions.Add(nodDisplayed.InnerText, True)
                        Next

                        NewState.htblObjectStates.Add(sKey, obs)
                    End With
                Next

                For Each nodTas As XmlElement In xmlDoc.SelectNodes("/Game/Task")
                    With nodTas
                        Dim tas As New clsGameState.clsTaskState
                        Dim sKey As String = .Item("Key").InnerText
                        tas.Completed = CBool(.Item("Completed").InnerText)
                        If .Item("Scored") IsNot Nothing Then tas.Scored = CBool(.Item("Scored").InnerText)
                        For Each nodDisplayed As XmlElement In nodTas.SelectNodes("Displayed")
                            tas.htblDisplayedDescriptions.Add(nodDisplayed.InnerText, True)
                        Next

                        NewState.htblTaskStates.Add(sKey, tas)
                    End With
                Next

                For Each nodEv As XmlElement In xmlDoc.SelectNodes("/Game/Event")
                    With nodEv
                        Dim evs As New clsGameState.clsEventState
                        Dim sKey As String = .Item("Key").InnerText
                        evs.Status = CType([Enum].Parse(GetType(clsEvent.StatusEnum), .Item("Status").InnerText), clsEvent.StatusEnum)
                        evs.TimerToEndOfEvent = SafeInt(.Item("Timer").InnerText)
                        If .Item("SubEventTime") IsNot Nothing Then evs.iLastSubEventTime = SafeInt(.Item("SubEventTime").InnerText)
                        If .Item("SubEventIndex") IsNot Nothing Then evs.iLastSubEventIndex = SafeInt(.Item("SubEventIndex").InnerText)
                        For Each nodDisplayed As XmlElement In nodEv.SelectNodes("Displayed")
                            evs.htblDisplayedDescriptions.Add(nodDisplayed.InnerText, True)
                        Next
                        NewState.htblEventStates.Add(sKey, evs)
                    End With
                Next

                For Each nodCh As XmlElement In xmlDoc.SelectNodes("/Game/Character")
                    With nodCh
                        Dim chs As New clsGameState.clsCharacterState
                        Dim sKey As String = .Item("Key").InnerText
                        If Adventure.htblCharacters.ContainsKey(sKey) Then
                            chs.Location = New clsCharacterLocation(Adventure.htblCharacters(sKey))
                            If .Item("ExistWhere") IsNot Nothing Then
                                chs.Location.ExistWhere = CType([Enum].Parse(GetType(clsCharacterLocation.ExistsWhereEnum), .Item("ExistWhere").InnerText), clsCharacterLocation.ExistsWhereEnum)
                            Else
                                chs.Location.ExistWhere = clsCharacterLocation.ExistsWhereEnum.Hidden
                            End If
                            If .Item("Position") IsNot Nothing Then
                                chs.Location.Position = CType([Enum].Parse(GetType(clsCharacterLocation.PositionEnum), .Item("Position").InnerText), clsCharacterLocation.PositionEnum)
                            Else
                                chs.Location.Position = clsCharacterLocation.PositionEnum.Standing
                            End If
                            If .Item("LocationKey") IsNot Nothing Then
                                chs.Location.Key = .Item("LocationKey").InnerText
                            Else
                                chs.Location.Key = ""
                            End If
                            For Each nodW As XmlElement In nodCh.SelectNodes("Walk")
                                With nodW
                                    Dim ws As New clsGameState.clsCharacterState.clsWalkState
                                    ws.Status = CType([Enum].Parse(GetType(clsWalk.StatusEnum), .Item("Status").InnerText), clsWalk.StatusEnum)
                                    ws.TimerToEndOfWalk = SafeInt(.Item("Timer").InnerText)
                                    chs.lWalks.Add(ws)
                                End With
                            Next
                            For Each nodProp As XmlElement In nodCh.SelectNodes("Property")
                                Dim props As New clsGameState.clsCharacterState.clsStateProperty
                                Dim sPropKey As String = nodProp.Item("Key").InnerText
                                props.Value = nodProp.Item("Value").InnerText
                                chs.htblProperties.Add(sPropKey, props)
                            Next
                            For Each nodSeen As XmlElement In nodCh.SelectNodes("Seen")
                                chs.lSeenKeys.Add(nodSeen.InnerText)
                            Next
                            For Each nodDisplayed As XmlElement In nodCh.SelectNodes("Displayed")
                                chs.htblDisplayedDescriptions.Add(nodDisplayed.InnerText, True)
                            Next
                            NewState.htblCharacterStates.Add(sKey, chs)
                        Else
                            DisplayError("Character key " & sKey & " not found in adventure!")
                        End If
                    End With
                Next

                For Each nodVar As XmlElement In xmlDoc.SelectNodes("/Game/Variable")
                    With nodVar
                        Dim vars As New clsGameState.clsVariableState
                        Dim sKey As String = .Item("Key").InnerText
                        If Adventure.htblVariables.ContainsKey(sKey) Then
                            Dim v As clsVariable = Adventure.htblVariables(sKey)

                            ReDim vars.Value(v.Length - 1)
                            For i As Integer = 0 To v.Length - 1
                                If .Item("Value_" & i) IsNot Nothing Then
                                    vars.Value(i) = SafeString(.Item("Value_" & i).InnerText)
                                Else
                                    If v.Type = clsVariable.VariableTypeEnum.Numeric Then
                                        vars.Value(i) = "0"
                                    Else
                                        vars.Value(i) = ""
                                    End If
                                End If
                            Next
                            If .Item("Value") IsNot Nothing Then ' Old style                            
                                vars.Value(0) = SafeString(.Item("Value").InnerText)
                            End If

                            For Each nodDisplayed As XmlElement In nodVar.SelectNodes("Displayed")
                                vars.htblDisplayedDescriptions.Add(nodDisplayed.InnerText, True)
                            Next
                            NewState.htblVariableStates.Add(sKey, vars)
                        Else
                            DisplayError("Variable key " & sKey & " not found in adventure!")
                        End If

                    End With
                Next

                For Each nodGrp As XmlElement In xmlDoc.SelectNodes("/Game/Group")
                    With nodGrp
                        Dim grps As New clsGameState.clsGroupState
                        Dim sKey As String = .Item("Key").InnerText
                        For Each nodMember As XmlElement In nodGrp.SelectNodes("Member")
                            grps.lstMembers.Add(nodMember.InnerText)
                        Next
                        NewState.htblGroupStates.Add(sKey, grps)
                    End With
                Next

                If .Item("Turns") IsNot Nothing Then Adventure.Turns = SafeInt(.Item("Turns").InnerText)

            End With

            Return NewState

        Catch ex As Exception
            ErrMsg("Error loading game state", ex)
        End Try

        Return Nothing

    End Function

    Private Sub Set400SpecificTasks()

        Dim tasGetParent As clsTask = Adventure.htblTasks("TakeObjects") ' Use the parent task, because we don't know if they're being lazy or not...

        If Not tasGetParent Is Nothing Then
            ' Convert get/drop tasks into Specific tasks - this may not be perfect, but it's a good guess                        
            For Each tas As clsTask In Adventure.htblTasks.Values
                For Each sCommand As String In tas.arlCommands
                    If sCommand.Contains("get ") OrElse sCommand = "get *" Then
                        ' We've got to work out if a particular object is referenced
                        ' First look to see if the command makes it unique
                        Dim iObsFound As Integer = 0
                        Dim sFound As String = ""
                        For Each ob As clsObject In Adventure.htblObjects.Values
                            If sCommand.Contains(ob.arlNames(0)) Then
                                If Not ob.IsStatic Then
                                    iObsFound += 1
                                    sFound = ob.Key
                                Else
                                    ' Don't override any tasks where the player is getting something and that something exists as a static object
                                    GoTo NextTask
                                End If
                            End If
                        Next
                        If iObsFound > 1 Then
                            ' Ok, lets look in the restrictions/actions to see if they point us to an exact object
                            For Each rest As clsRestriction In tas.arlRestrictions
                                If Adventure.htblObjects.ContainsKey(rest.sKey1) AndAlso sCommand.Contains(Adventure.htblObjects(rest.sKey1).arlNames(0)) Then
                                    sFound = rest.sKey1
                                ElseIf Adventure.htblObjects.ContainsKey(rest.sKey2) AndAlso sCommand.Contains(Adventure.htblObjects(rest.sKey2).arlNames(0)) Then
                                    sFound = rest.sKey2
                                End If
                                If sFound <> "" Then Exit For
                            Next
                            If sFound = "" Then
                                For Each act As clsAction In tas.arlActions
                                    If Adventure.htblObjects.ContainsKey(act.sKey1) AndAlso sCommand.Contains(Adventure.htblObjects(act.sKey1).arlNames(0)) Then
                                        sFound = act.sKey1
                                    ElseIf Adventure.htblObjects.ContainsKey(act.sKey2) AndAlso sCommand.Contains(Adventure.htblObjects(act.sKey2).arlNames(0)) Then
                                        sFound = act.sKey2
                                    End If
                                    If sFound <> "" Then Exit For
                                Next
                            End If
                        End If
                        If sFound <> "" OrElse sCommand = "get *" Then
                            tas.TaskType = clsTask.TaskTypeEnum.Specific
                            tas.GeneralKey = tasGetParent.Key
                            ReDim tas.Specifics(0)
                            Dim s As New clsTask.Specific
                            With s
                                .Multiple = False
                                .Keys.Add(sFound)
                            End With
                            tas.Specifics(0) = s
                            Exit For
                        End If
                    End If
                Next
NextTask:
            Next
        End If

        Dim tasExamineChar As clsTask = Adventure.htblTasks("ExamineCharacter") ' Use the parent task, because we don't know if they're being lazy or not...
        If tasExamineChar IsNot Nothing Then
            ' Make it a system task, i.e. don't run events
            tasExamineChar.bSystemTask = True
        End If

    End Sub



    Private Sub ObfuscateByteArray(ByRef bytData As Byte(), Optional ByVal iOffset As Integer = 0, Optional ByVal iLength As Integer = 0)

        Dim iRandomKey As Integer() = {41, 236, 221, 117, 23, 189, 44, 187, 161, 96, 4, 147, 90, 91, 172, 159, 244, 50, 249, 140, 190, 244, 82, 111, 170, 217, 13, 207, 25, 177, 18, 4, 3, 221, 160, 209, 253, 69, 131, 37, 132, 244, 21, 4, 39, 87, 56, 203, 119, 139, 231, 180, 190, 13, 213, 53, 153, 109, 202, 62, 175, 93, 161, 239, 77, 0, 143, 124, 186, 219, 161, 175, 175, 212, 7, 202, 223, 77, 72, 83, 160, 66, 88, 142, 202, 93, 70, 246, 8, 107, 55, 144, 122, 68, 117, 39, 83, 37, 183, 39, 199, 188, 16, 155, 233, 55, 5, 234, 6, 11, 86, 76, 36, 118, 158, 109, 5, 19, 36, 239, 185, 153, 115, 79, 164, 17, 52, 106, 94, 224, 118, 185, 150, 33, 139, 228, 49, 188, 164, 146, 88, 91, 240, 253, 21, 234, 107, 3, 166, 7, 33, 63, 0, 199, 109, 46, 72, 193, 246, 216, 3, 154, 139, 37, 148, 156, 182, 3, 235, 185, 60, 73, 111, 145, 151, 94, 169, 118, 57, 186, 165, 48, 195, 86, 190, 55, 184, 206, 180, 93, 155, 111, 197, 203, 143, 189, 208, 202, 105, 121, 51, 104, 24, 237, 203, 216, 208, 111, 48, 15, 132, 210, 136, 60, 51, 211, 215, 52, 102, 92, 227, 232, 79, 142, 29, 204, 131, 163, 2, 217, 141, 223, 12, 192, 134, 61, 23, 214, 139, 230, 102, 73, 158, 165, 216, 201, 231, 137, 152, 187, 230, 155, 99, 12, 149, 75, 25, 138, 207, 254, 85, 44, 108, 86, 129, 165, 197, 200, 182, 245, 187, 1, 169, 128, 245, 153, 74, 170, 181, 83, 229, 250, 11, 70, 243, 242, 123, 0, 42, 58, 35, 141, 6, 140, 145, 58, 221, 71, 35, 51, 4, 30, 210, 162, 0, 229, 241, 227, 22, 252, 1, 110, 212, 123, 24, 90, 32, 37, 99, 142, 42, 196, 158, 123, 209, 45, 250, 28, 238, 187, 188, 3, 134, 130, 79, 199, 39, 105, 70, 14, 0, 151, 234, 46, 56, 181, 185, 138, 115, 54, 25, 183, 227, 149, 9, 63, 128, 87, 208, 210, 234, 213, 244, 91, 63, 254, 232, 81, 44, 81, 51, 183, 222, 85, 142, 146, 218, 112, 66, 28, 116, 111, 168, 184, 161, 4, 31, 241, 121, 15, 70, 208, 152, 116, 35, 43, 163, 142, 238, 58, 204, 103, 94, 34, 2, 97, 217, 142, 6, 119, 100, 16, 20, 179, 94, 122, 44, 59, 185, 58, 223, 247, 216, 28, 11, 99, 31, 105, 49, 98, 238, 75, 129, 8, 80, 12, 17, 134, 181, 63, 43, 145, 234, 2, 170, 54, 188, 228, 22, 168, 255, 103, 213, 180, 91, 213, 143, 65, 23, 159, 66, 111, 92, 164, 136, 25, 143, 11, 99, 81, 105, 165, 133, 121, 14, 77, 12, 213, 114, 213, 166, 58, 83, 136, 99, 135, 118, 205, 173, 123, 124, 207, 111, 22, 253, 188, 52, 70, 122, 145, 167, 176, 129, 196, 63, 89, 225, 91, 165, 13, 200, 185, 207, 65, 248, 8, 27, 211, 64, 1, 162, 193, 94, 231, 213, 153, 53, 111, 124, 81, 25, 198, 91, 224, 45, 246, 184, 142, 73, 9, 165, 26, 39, 159, 178, 194, 0, 45, 29, 245, 161, 97, 5, 120, 238, 229, 81, 153, 239, 165, 35, 114, 223, 83, 244, 1, 94, 238, 20, 2, 79, 140, 137, 54, 91, 136, 153, 190, 53, 18, 153, 8, 81, 135, 176, 184, 193, 226, 242, 72, 164, 30, 159, 164, 230, 51, 58, 212, 171, 176, 100, 17, 25, 27, 165, 20, 215, 206, 29, 102, 75, 147, 100, 221, 11, 27, 32, 88, 162, 59, 64, 123, 252, 203, 93, 48, 237, 229, 80, 40, 77, 197, 18, 132, 173, 136, 238, 54, 225, 156, 225, 242, 197, 140, 252, 17, 185, 193, 153, 202, 19, 226, 49, 112, 111, 232, 20, 78, 190, 117, 38, 242, 125, 244, 24, 134, 128, 224, 47, 130, 45, 234, 119, 6, 90, 78, 182, 112, 206, 76, 118, 43, 75, 134, 20, 107, 147, 162, 20, 197, 116, 160, 119, 107, 117, 238, 116, 208, 115, 118, 144, 217, 146, 22, 156, 41, 107, 43, 21, 33, 50, 163, 127, 114, 254, 251, 166, 247, 223, 173, 242, 222, 203, 106, 14, 141, 114, 11, 145, 107, 217, 229, 253, 88, 187, 156, 153, 53, 233, 235, 255, 104, 141, 243, 146, 209, 33, 5, 109, 122, 72, 125, 240, 198, 131, 178, 14, 40, 8, 15, 182, 95, 153, 169, 71, 77, 166, 38, 182, 97, 97, 113, 13, 244, 173, 138, 80, 215, 215, 61, 107, 108, 157, 22, 35, 91, 244, 55, 213, 8, 142, 113, 44, 217, 52, 159, 206, 228, 171, 68, 42, 250, 78, 11, 24, 215, 112, 252, 24, 249, 97, 54, 80, 202, 164, 74, 194, 131, 133, 235, 88, 110, 81, 173, 211, 240, 68, 51, 191, 13, 187, 108, 44, 147, 18, 113, 30, 146, 253, 76, 235, 247, 30, 219, 167, 88, 32, 97, 53, 234, 221, 75, 94, 192, 236, 188, 169, 160, 56, 40, 146, 60, 61, 10, 62, 245, 10, 189, 184, 50, 43, 47, 133, 57, 0, 97, 80, 117, 6, 122, 207, 226, 253, 212, 48, 112, 14, 108, 166, 86, 199, 125, 89, 213, 185, 174, 186, 20, 157, 178, 78, 99, 169, 2, 191, 173, 197, 36, 191, 139, 107, 52, 154, 190, 88, 175, 63, 105, 218, 206, 230, 157, 22, 98, 107, 174, 214, 175, 127, 81, 166, 60, 215, 84, 44, 107, 57, 251, 21, 130, 170, 233, 172, 27, 234, 147, 227, 155, 125, 10, 111, 80, 57, 207, 203, 176, 77, 71, 151, 16, 215, 22, 165, 110, 228, 47, 92, 69, 145, 236, 118, 68, 84, 88, 35, 252, 241, 250, 119, 215, 203, 59, 50, 117, 225, 86, 2, 8, 137, 124, 30, 242, 99, 4, 171, 148, 68, 61, 55, 186, 55, 157, 9, 144, 147, 43, 252, 225, 171, 206, 190, 83, 207, 191, 68, 155, 227, 47, 140, 142, 45, 84, 188, 20}

        For i As Integer = 0 To bytData.Length - 1
            If i >= iOffset AndAlso (iLength = 0 OrElse i < iLength + iOffset) Then bytData(i) = CByte(bytData(i) Xor iRandomKey((i - iOffset) Mod 1024))
        Next

    End Sub


    Private Function Decompress(ByVal bZLib As Byte(), ByVal bObfuscate As Boolean, Optional ByVal iOffset As Integer = 0, Optional ByVal iLength As Integer = 0) As MemoryStream

        If bObfuscate Then ObfuscateByteArray(bZLib, iOffset, iLength)
        Dim outStream As New System.IO.MemoryStream
        If iLength = 0 Then iLength = bZLib.Length - iOffset
        Dim inStream As New System.IO.MemoryStream(bZLib, iOffset, iLength)
        Dim zStream As New InflaterInputStream(inStream)
        Try
            If Not CopyStream(zStream, outStream) Then Return Nothing
            Return New MemoryStream(outStream.GetBuffer)
        Catch ex As Exception
            ErrMsg("Error decompressing byte array", ex)
        Finally
            zStream.Close()
            outStream.Close()
            inStream.Close()
        End Try

        Return Nothing
    End Function

    Private Function FileToMemoryStream(ByVal bCompressed As Boolean, ByVal iLength As Integer, ByVal bObfuscate As Boolean) As MemoryStream

        'ReDim bAdventure(-1) ' if this needs to be done, why here?

        If bCompressed Then
            Dim bAdvZLib() As Byte = br.ReadBytes(iLength)
            Return Decompress(bAdvZLib, bObfuscate)
        Else
            Return New MemoryStream(br.ReadBytes(iLength))
        End If

    End Function


    ' Grab the numeric part of the key (if any) and increment it
    Private Function IncrementKey(ByVal sKey As String) As String

        Dim re As New System.Text.RegularExpressions.Regex("")

        Dim sJustKey As String = System.Text.RegularExpressions.Regex.Replace(sKey, "\d*$", "")
        Dim sNumber As String = sKey.Replace(sJustKey, "")
        If sNumber = "" Then sNumber = "0"
        Dim iNumber As Integer = CInt(sNumber)
        Return sJustKey & iNumber + 1

    End Function


    Private Class LoadAbortException
        Inherits Exception
    End Class

    Private Enum LoadItemEum As Integer
        No = 0
        Yes = 1
        Both = 2
    End Enum
    Private Function ShouldWeLoadLibraryItem(ByVal sItemKey As String) As LoadItemEum
        If Adventure.listExcludedItems.Contains(sItemKey) Then Return LoadItemEum.No
        Return LoadItemEum.Yes
    End Function


    Private Sub SetChar(ByRef sText As String, ByVal iPos As Integer, ByVal Character As Char)
        If iPos > sText.Length - 1 Then Exit Sub
        sText = sLeft(sText, iPos) & Character & sRight(sText, sText.Length - iPos - 1)
    End Sub


    Private Function GetDate(ByVal sDate As String) As Date

        ' Non UK locals save "yyyy-MM-dd HH:mm:ss" in their own format.  Convert these for pre-5.0.22 saves
        If dFileVersion < 5.000022 Then
            If sDate.Length = 19 Then
                SetChar(sDate, 13, ":"c)
                SetChar(sDate, 16, ":"c)
            End If
        End If

        Dim dtReturn As Date
        If Date.TryParse(sDate, dtReturn) Then Return dtReturn

        Return Date.MinValue

    End Function


    Private Function SetDate(ByVal dtDate As Date) As String
        Return dtDate.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat)
    End Function


    ' Corrects the bracket sequences for ORs after ANDs
    Private Function CorrectBracketSequence(ByVal sSequence As String) As String

        If sSequence.Contains("#A#O#") Then
            For i As Integer = 10 To 0 Step -1
                Dim sSearch As String = "#A#"
                For j As Integer = 0 To i
                    sSearch &= "O#"
                Next
                While sSequence.Contains(sSearch)
                    Dim sReplace As String = "#A(#"
                    For j As Integer = 0 To i
                        sReplace &= "O#"
                    Next
                    sReplace &= ")"
                    sSequence = sSequence.Replace(sSearch, sReplace)
                    iCorrectedTasks += 1
                End While
            Next
        End If

        Return sSequence

    End Function


    Private Function GetBool(ByVal sBool As String) As Boolean

        Select Case sBool.ToUpper
            Case "0", "FALSE"
                Return False
            Case "-1", "1", "TRUE", "VRAI"
                Return True
            Case Else
                Return False
        End Select

    End Function


    Dim bCorrectBracketSequences As Boolean = False
    Dim bAskedAboutBrackets As Boolean = False
    Dim iCorrectedTasks As Integer = 0

    Private Function Load500(ByVal stmMemory As MemoryStream, ByVal bLibrary As Boolean, Optional ByVal bAppend As Boolean = False, Optional ByVal eLoadWhat As LoadWhatEnum = LoadWhatEnum.All, Optional ByVal dtAdvDate As Date = #1/1/1900#, Optional ByVal sFilename As String = "") As Boolean
        Try
            If stmMemory Is Nothing Then Return False

            Dim xmlDoc As New XmlDocument
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(stmMemory)
            Dim a As clsAdventure = Adventure
            Dim bAddDuplicateKeys As Boolean = True
            Dim htblDuplicateKeyMapping As New StringHashTable
            Dim arlNewTasks As New StringArrayList

            With xmlDoc.Item("Adventure")
                If .Item("Version") IsNot Nothing Then
                    dFileVersion = SafeDbl(.Item("Version").InnerText)
                    If dFileVersion > dVersion Then
                        Glue.MakeNote("This file is newer than what this software might understand.")
                    End If
                    If Not bLibrary Then a.dVersion = dFileVersion
                Else
                    Throw New Exception("Version tag not specified")
                End If

                If eLoadWhat = LoadWhatEnum.All Then
                    bAskedAboutBrackets = False
                    iCorrectedTasks = 0

                    If Not bLibrary AndAlso .Item("Title") IsNot Nothing Then a.Title = .Item("Title").InnerText
                    If Not bLibrary AndAlso .Item("Author") IsNot Nothing Then a.Author = .Item("Author").InnerText
                    If .Item("LastUpdated") IsNot Nothing Then a.LastUpdated = GetDate(.Item("LastUpdated").InnerText)
                    If Not .Item("Introduction") Is Nothing Then a.Introduction = LoadDescription(xmlDoc.Item("Adventure"), "Introduction") ' New Description(.Item("Introduction").InnerText)
                    If .Item("FontName") IsNot Nothing Then a.DefaultFontName = .Item("FontName").InnerText
                    If .Item("FontSize") IsNot Nothing Then a.DefaultFontSize = SafeInt(.Item("FontSize").InnerText)
                    a.DeveloperDefaultBackgroundColour = Nothing
                    a.DeveloperDefaultInputColour = Nothing
                    a.DeveloperDefaultOutputColour = Nothing
                    a.DeveloperDefaultLinkColour = Nothing
                    If .Item("BackgroundColour") IsNot Nothing Then a.DeveloperDefaultBackgroundColour = System.Drawing.ColorTranslator.FromOle(CInt(.Item("BackgroundColour").InnerText))
                    If .Item("InputColour") IsNot Nothing Then a.DeveloperDefaultInputColour = System.Drawing.ColorTranslator.FromOle(CInt(.Item("InputColour").InnerText))
                    If .Item("OutputColour") IsNot Nothing Then a.DeveloperDefaultOutputColour = System.Drawing.ColorTranslator.FromOle(CInt(.Item("OutputColour").InnerText))
                    If .Item("LinkColour") IsNot Nothing Then a.DeveloperDefaultLinkColour = System.Drawing.ColorTranslator.FromOle(CInt(.Item("LinkColour").InnerText))
                    If .Item("ShowFirstLocation") IsNot Nothing Then a.ShowFirstRoom = GetBool(.Item("ShowFirstLocation").InnerText)
                    If .Item("UserStatus") IsNot Nothing Then a.sUserStatus = .Item("UserStatus").InnerText
                    If .Item("ifindex") IsNot Nothing Then Adventure.BabelTreatyInfo.FromString(.Item("ifindex").OuterXml) ' Pre 5.0.20
                    If .Item("Cover") IsNot Nothing Then
                        Adventure.CoverFilename = .Item("Cover").InnerText
                    End If
                    If .Item("ShowExits") IsNot Nothing Then a.ShowExits = GetBool(.Item("ShowExits").InnerText)
                    If .Item("EnableMenu") IsNot Nothing Then a.EnableMenu = GetBool(.Item("EnableMenu").InnerText)
                    If .Item("EnableDebugger") IsNot Nothing Then a.EnableDebugger = GetBool(.Item("EnableDebugger").InnerText)
                    If .Item("EndGameText") IsNot Nothing Then a.WinningText = LoadDescription(xmlDoc.Item("Adventure"), "EndGameText")
                    If .Item("Elapsed") IsNot Nothing Then a.iElapsed = SafeInt(.Item("Elapsed").InnerText)
                    If .Item("TaskExecution") IsNot Nothing Then a.TaskExecution = CType([Enum].Parse(GetType(clsAdventure.TaskExecutionEnum), .Item("TaskExecution").InnerText), clsAdventure.TaskExecutionEnum)
                    If .Item("WaitTurns") IsNot Nothing Then a.WaitTurns = SafeInt(.Item("WaitTurns").InnerText)
                    If .Item("KeyPrefix") IsNot Nothing Then
                        a.KeyPrefix = .Item("KeyPrefix").InnerText
                    End If

                    If .Item("DirectionNorth") IsNot Nothing AndAlso .Item("DirectionNorth").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.North) = .Item("DirectionNorth").InnerText
                    If .Item("DirectionNorthEast") IsNot Nothing AndAlso .Item("DirectionNorthEast").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.NorthEast) = .Item("DirectionNorthEast").InnerText
                    If .Item("DirectionEast") IsNot Nothing AndAlso .Item("DirectionEast").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.East) = .Item("DirectionEast").InnerText
                    If .Item("DirectionSouthEast") IsNot Nothing AndAlso .Item("DirectionSouthEast").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.SouthEast) = .Item("DirectionSouthEast").InnerText
                    If .Item("DirectionSouth") IsNot Nothing AndAlso .Item("DirectionSouth").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.South) = .Item("DirectionSouth").InnerText
                    If .Item("DirectionSouthWest") IsNot Nothing AndAlso .Item("DirectionSouthWest").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.SouthWest) = .Item("DirectionSouthWest").InnerText
                    If .Item("DirectionWest") IsNot Nothing AndAlso .Item("DirectionWest").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.West) = .Item("DirectionWest").InnerText
                    If .Item("DirectionNorthWest") IsNot Nothing AndAlso .Item("DirectionNorthWest").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.NorthWest) = .Item("DirectionNorthWest").InnerText
                    If .Item("DirectionIn") IsNot Nothing AndAlso .Item("DirectionIn").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.In) = .Item("DirectionIn").InnerText
                    If .Item("DirectionOut") IsNot Nothing AndAlso .Item("DirectionOut").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.Out) = .Item("DirectionOut").InnerText
                    If .Item("DirectionUp") IsNot Nothing AndAlso .Item("DirectionUp").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.Up) = .Item("DirectionUp").InnerText
                    If .Item("DirectionDown") IsNot Nothing AndAlso .Item("DirectionDown").InnerText <> "" Then a.sDirectionsRE(DirectionsEnum.Down) = .Item("DirectionDown").InnerText

                End If

                If eLoadWhat = LoadWhatEnum.All OrElse eLoadWhat = LoadWhatEnum.AllExceptProperties Then
                    Debug.WriteLine("End Intro: " & Now)
                End If


                If eLoadWhat = LoadWhatEnum.All OrElse eLoadWhat = LoadWhatEnum.Properies Then
                    ' Properties
                    For Each nodProp As XmlElement In xmlDoc.SelectNodes("/Adventure/Property")
                        Dim prop As New clsProperty
                        With nodProp
                            Dim sKey As String = .Item("Key").InnerText
                            If .Item("Library") IsNot Nothing Then prop.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblAllProperties.ContainsKey(sKey) Then
                                If a.htblAllProperties.ContainsKey(sKey) Then
                                    If prop.IsLibrary OrElse bLibrary Then
                                        If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblAllProperties(sKey).LastUpdated Then GoTo NextProp
                                        Select Case ShouldWeLoadLibraryItem(sKey)
                                            Case LoadItemEum.Yes
                                                a.htblAllProperties.Remove(sKey)
                                            Case LoadItemEum.No
                                                GoTo NextProp
                                            Case LoadItemEum.Both
                                                ' Keep key, but still add this new one
                                        End Select
                                    End If
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblAllProperties.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextProp
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextProp
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            prop.Key = sKey
                            If bLibrary Then
                                prop.IsLibrary = True
                                prop.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then prop.LastUpdated = GetDate(.Item("LastUpdated").InnerText)
                            prop.Description = .Item("Description").InnerText
                            If Not .Item("Mandatory") Is Nothing Then
                                prop.Mandatory = GetBool(.Item("Mandatory").InnerText)
                            End If
                            If .Item("PropertyOf") IsNot Nothing Then prop.PropertyOf = EnumParsePropertyPropertyOf(.Item("PropertyOf").InnerText)
                            If .Item("AppendTo") IsNot Nothing Then prop.AppendToProperty = .Item("AppendTo").InnerText
                            prop.Type = EnumParsePropertyType(.Item("Type").InnerText)
                            Select Case prop.Type
                                Case clsProperty.PropertyTypeEnum.StateList
                                    For Each nodState As XmlElement In nodProp.SelectNodes("State")
                                        prop.arlStates.Add(nodState.InnerText)
                                    Next
                                    If prop.arlStates.Count > 0 Then prop.Value = prop.arlStates(0)
                                Case clsProperty.PropertyTypeEnum.ValueList
                                    For Each nodValueList As XmlElement In nodProp.SelectNodes("ValueList")
                                        If nodValueList.Item("Label") IsNot Nothing Then
                                            Dim sLabel As String = nodValueList("Label").InnerText
                                            Dim iValue As Integer = 0
                                            If nodValueList.Item("Value") IsNot Nothing Then iValue = SafeInt(nodValueList("Value").InnerText)
                                            prop.ValueList.Add(sLabel, iValue)
                                        End If
                                    Next
                            End Select
                            If .Item("PrivateTo") IsNot Nothing Then prop.PrivateTo = .Item("PrivateTo").InnerText
                            If .Item("Tooltip") IsNot Nothing Then prop.PopupDescription = .Item("Tooltip").InnerText

                            If Not .Item("DependentKey") Is Nothing Then
                                If .Item("DependentKey").InnerText <> sKey Then
                                    prop.DependentKey = .Item("DependentKey").InnerText
                                    If Not .Item("DependentValue") Is Nothing Then
                                        prop.DependentValue = .Item("DependentValue").InnerText
                                    End If
                                End If
                            End If
                            If Not .Item("RestrictProperty") Is Nothing Then
                                prop.RestrictProperty = .Item("RestrictProperty").InnerText
                                If Not .Item("RestrictValue") Is Nothing Then
                                    prop.RestrictValue = .Item("RestrictValue").InnerText
                                End If
                            End If
                        End With
                        a.htblAllProperties.Add(prop)
NextProp:
                    Next
                    a.htblAllProperties.SetSelected()
                    Debug.WriteLine("End Properties: " & Now)

                    CreateMandatoryProperties()
                End If



                If eLoadWhat = LoadWhatEnum.All OrElse eLoadWhat = LoadWhatEnum.AllExceptProperties Then
                    ' Locations
                    For Each nodLoc As XmlElement In xmlDoc.SelectNodes("/Adventure/Location")
                        Dim loc As New clsLocation
                        With nodLoc
                            Dim sKey As String = .Item("Key").InnerText
                            If Not .Item("Library") Is Nothing Then loc.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblLocations.ContainsKey(sKey) Then
                                If a.htblLocations.ContainsKey(sKey) Then
                                    If loc.IsLibrary OrElse bLibrary Then
                                        If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblLocations(sKey).LastUpdated Then GoTo NextLoc
                                        Select Case ShouldWeLoadLibraryItem(sKey)
                                            Case LoadItemEum.Yes
                                                a.htblLocations.Remove(sKey)
                                            Case LoadItemEum.No
                                                GoTo NextLoc
                                            Case LoadItemEum.Both
                                                ' Keep key, but still add this new one
                                        End Select
                                    End If
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblLocations.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextLoc
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextLoc
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            loc.Key = sKey
                            If bLibrary Then
                                loc.IsLibrary = True
                                loc.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then loc.LastUpdated = CDate(.Item("LastUpdated").InnerText)
                            If dFileVersion < 5.000015 Then
                                loc.ShortDescription = New Description(.Item("ShortDescription").InnerText)
                            Else
                                loc.ShortDescription = LoadDescription(nodLoc, "ShortDescription")
                            End If
                            loc.LongDescription = LoadDescription(nodLoc, "LongDescription")

                            For Each nodDir As XmlElement In nodLoc.SelectNodes("Movement")
                                Dim xdir As DirectionsEnum = EnumParseDirections(nodDir.Item("Direction").InnerText)

                                With loc.arlDirections(xdir)
                                    .LocationKey = nodDir.Item("Destination").InnerText
                                    .Restrictions = LoadRestrictions(nodDir)
                                End With
                            Next
                            For Each nodProp As XmlElement In .SelectNodes("Property")
                                Dim prop As New clsProperty
                                Dim sPropKey As String = nodProp.Item("Key").InnerText
                                If Adventure.htblAllProperties.ContainsKey(sPropKey) Then
                                    prop = Adventure.htblAllProperties(sPropKey).Copy
                                    If prop.Type = clsProperty.PropertyTypeEnum.Text Then
                                        prop.StringData = LoadDescription(nodProp, "Value")
                                    ElseIf prop.Type <> clsProperty.PropertyTypeEnum.SelectionOnly Then
                                        prop.Value = nodProp.Item("Value").InnerText
                                    End If
                                    prop.Selected = True
                                    loc.AddProperty(prop)
                                End If
                            Next
                            If .Item("Hide") IsNot Nothing Then loc.HideOnMap = GetBool(.Item("Hide").InnerText)
                        End With
                        a.htblLocations.Add(loc, loc.Key)
NextLoc:
                    Next nodLoc
                    Debug.WriteLine("End Locations: " & Now)
                End If



                If eLoadWhat = LoadWhatEnum.All OrElse eLoadWhat = LoadWhatEnum.AllExceptProperties Then
                    ' Objects
                    For Each nodOb As XmlElement In .SelectNodes("/Adventure/Object")
                        Dim ob As New clsObject
                        With nodOb
                            Dim sKey As String = .Item("Key").InnerText
                            If Not .Item("Library") Is Nothing Then ob.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblObjects.ContainsKey(sKey) Then
                                If a.htblObjects.ContainsKey(sKey) Then
                                    If ob.IsLibrary OrElse bLibrary Then
                                        If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblObjects(sKey).LastUpdated Then GoTo NextOb
                                        Select Case ShouldWeLoadLibraryItem(sKey)
                                            Case LoadItemEum.Yes
                                                a.htblObjects.Remove(sKey)
                                            Case LoadItemEum.No
                                                GoTo NextOb
                                            Case LoadItemEum.Both
                                                ' Keep key, but still add this new one
                                        End Select
                                    End If
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblObjects.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextOb
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextOb
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            ob.Key = sKey
                            If bLibrary Then
                                ob.IsLibrary = True
                                ob.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then ob.LastUpdated = CDate(.Item("LastUpdated").InnerText)
                            If .Item("Article") IsNot Nothing Then ob.Article = .Item("Article").InnerText
                            If .Item("Prefix") IsNot Nothing Then ob.Prefix = .Item("Prefix").InnerText
                            For Each nodName As XmlElement In .GetElementsByTagName("Name")
                                ob.arlNames.Add(nodName.InnerText)
                            Next
                            ob.Description = LoadDescription(nodOb, "Description")

                            For Each nodProp As XmlElement In .SelectNodes("Property")
                                Dim prop As New clsProperty
                                Dim sPropKey As String = nodProp.Item("Key").InnerText
                                If Adventure.htblAllProperties.ContainsKey(sPropKey) Then
                                    prop = Adventure.htblAllProperties(sPropKey).Copy
                                    If prop.Type = clsProperty.PropertyTypeEnum.Text Then
                                        prop.StringData = LoadDescription(nodProp, "Value")
                                    ElseIf prop.Type <> clsProperty.PropertyTypeEnum.SelectionOnly Then
                                        prop.Value = nodProp.Item("Value").InnerText
                                    End If
                                    prop.Selected = True
                                    ob.AddProperty(prop)
                                End If
                            Next
                            ob.Location = ob.Location ' Assigns the location object from the object properties                            
                        End With
                        a.htblObjects.Add(ob, ob.Key)
NextOb:
                    Next
                    Debug.WriteLine("End Objects: " & Now)

                    ' Tasks                    
                    For Each nodTask As XmlElement In .SelectNodes("/Adventure/Task")
                        Dim tas As New clsTask
                        With nodTask
                            Dim sKey As String = .Item("Key").InnerText
                            If .Item("Library") IsNot Nothing Then tas.IsLibrary = GetBool(.Item("Library").InnerText)
                            If .Item("ReplaceTask") IsNot Nothing Then tas.ReplaceDuplicateKey = GetBool(.Item("ReplaceTask").InnerText)
                            If a.htblTasks.ContainsKey(sKey) Then
                                If tas.IsLibrary OrElse bLibrary Then
                                    ' We skip loading the task if it is not newer than the one we currently have loaded
                                    ' If there's no timestamp, we have to assume it is newer
                                    If .Item("LastUpdated") IsNot Nothing AndAlso CDate(.Item("LastUpdated").InnerText) <= a.htblTasks(sKey).LastUpdated Then GoTo NextTask
                                    ' If a library item is newer than the library in your game, prompt
                                    Select Case ShouldWeLoadLibraryItem(sKey)
                                        Case LoadItemEum.Yes
                                            a.htblTasks.Remove(sKey)
                                        Case LoadItemEum.No
                                            ' Set the timestamp of the custom version to now, so it's more recent than the "newer" library.  That way we won't be prompted next time
                                            a.htblTasks(sKey).LastUpdated = Now
                                            GoTo NextTask
                                        Case LoadItemEum.Both
                                            ' Keep key, but still add this new one
                                    End Select
                                End If
                                If tas.ReplaceDuplicateKey Then
                                    If a.htblTasks.ContainsKey(sKey) Then a.htblTasks.Remove(sKey)
                                Else
                                    If bAddDuplicateKeys Then
                                        Dim sOldKey As String = sKey
                                        While a.htblTasks.ContainsKey(sKey)
                                            sKey = IncrementKey(sKey)
                                        End While
                                        htblDuplicateKeyMapping.Add(sOldKey, sKey)
                                    Else
                                        GoTo NextTask
                                    End If
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextTask
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            tas.Key = sKey
                            tas.Priority = CInt(.Item("Priority").InnerText)
                            If bLibrary AndAlso Not tas.IsLibrary Then tas.Priority += 50000
                            If bLibrary Then
                                tas.IsLibrary = True
                                tas.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then tas.LastUpdated = GetDate(.Item("LastUpdated").InnerText)
                            tas.Priority = CInt(.Item("Priority").InnerText)
                            If Not .Item("AutoFillPriority") Is Nothing Then tas.AutoFillPriority = CInt(.Item("AutoFillPriority").InnerText)
                            tas.TaskType = EnumParseTaskType(.Item("Type").InnerText)
                            tas.CompletionMessage = LoadDescription(nodTask, "CompletionMessage")
                            Select Case tas.TaskType
                                Case clsTask.TaskTypeEnum.General
                                    For Each nodCommand As XmlElement In .GetElementsByTagName("Command")
                                        ' Simplify Runner so it only has to deal with multiple, or specific refs
                                        tas.arlCommands.Add(FixInitialRefs(nodCommand.InnerText))
                                    Next
                                Case clsTask.TaskTypeEnum.Specific
                                    tas.GeneralKey = .Item("GeneralTask").InnerText
                                    Dim iSpecCount As Integer = 0
                                    ReDim tas.Specifics(-1)
                                    For Each nodSpec As XmlElement In .GetElementsByTagName("Specific")
                                        Dim spec As New clsTask.Specific
                                        iSpecCount += 1
                                        spec.Type = EnumParseSpecificType(nodSpec.Item("Type").InnerText)
                                        spec.Multiple = GetBool(nodSpec.Item("Multiple").InnerText)
                                        For Each nodKey As XmlElement In nodSpec.GetElementsByTagName("Key")
                                            spec.Keys.Add(nodKey.InnerText)
                                        Next
                                        ReDim Preserve tas.Specifics(iSpecCount - 1)
                                        tas.Specifics(iSpecCount - 1) = spec
                                    Next
                                    If .Item("ExecuteParentActions") IsNot Nothing Then ' Old checkbox method
                                        If GetBool(.Item("ExecuteParentActions").InnerText) Then
                                            If tas.CompletionMessage.ToString(True) = "" Then
                                                tas.SpecificOverrideType = clsTask.SpecificOverrideTypeEnum.BeforeTextAndActions
                                            Else
                                                tas.SpecificOverrideType = clsTask.SpecificOverrideTypeEnum.BeforeActionsOnly
                                            End If
                                        Else
                                            If tas.CompletionMessage.ToString(True) = "" Then
                                                tas.SpecificOverrideType = clsTask.SpecificOverrideTypeEnum.BeforeTextOnly
                                            Else
                                                tas.SpecificOverrideType = clsTask.SpecificOverrideTypeEnum.Override
                                            End If
                                        End If
                                    End If
                                    If .Item("SpecificOverrideType") IsNot Nothing Then
                                        tas.SpecificOverrideType = CType([Enum].Parse(GetType(clsTask.SpecificOverrideTypeEnum), .Item("SpecificOverrideType").InnerText), clsTask.SpecificOverrideTypeEnum)
                                    End If
                                Case clsTask.TaskTypeEnum.System
                                    If .Item("RunImmediately") IsNot Nothing Then tas.RunImmediately = GetBool(.Item("RunImmediately").InnerText)
                                    If .Item("LocationTrigger") IsNot Nothing Then tas.LocationTrigger = .Item("LocationTrigger").InnerText
                            End Select
                            tas.Description = .Item("Description").InnerText
                            tas.Repeatable = GetBool(.Item("Repeatable").InnerText)
                            If .Item("Aggregate") IsNot Nothing Then tas.AggregateOutput = GetBool(.Item("Aggregate").InnerText)
                            If .Item("Continue") IsNot Nothing Then
                                Select Case .Item("Continue").InnerText
                                    Case "ContinueNever", "ContinueOnFail", "ContinueOnNoOutput"
                                        tas.ContinueToExecuteLowerPriority = False
                                    Case "ContinueAlways"
                                        tas.ContinueToExecuteLowerPriority = True
                                End Select
                            End If
                            If .Item("LowPriority") IsNot Nothing Then tas.LowPriority = GetBool(.Item("LowPriority").InnerText)
                            tas.arlRestrictions = LoadRestrictions(nodTask)
                            tas.arlActions = LoadActions(nodTask)
                            tas.FailOverride = LoadDescription(nodTask, "FailOverride")
                            If Not .Item("MessageBeforeOrAfter") Is Nothing Then tas.eDisplayCompletion = EnumParseBeforeAfter(.Item("MessageBeforeOrAfter").InnerText) Else tas.eDisplayCompletion = clsTask.BeforeAfterEnum.Before
                            If .Item("PreventOverriding") IsNot Nothing Then tas.PreventOverriding = GetBool(.Item("PreventOverriding").InnerText)
                        End With
                        a.htblTasks.Add(tas, tas.Key)
                        arlNewTasks.Add(tas.Key)
NextTask:
                    Next
                    Debug.WriteLine("End Tasks: " & Now)


                    For Each nodEvent As XmlElement In .SelectNodes("/Adventure/Event")
                        Dim ev As New clsEvent
                        With nodEvent
                            Dim sKey As String = .Item("Key").InnerText
                            If .Item("Library") IsNot Nothing Then ev.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblEvents.ContainsKey(sKey) Then
                                If a.htblEvents.ContainsKey(sKey) Then
                                    If ev.IsLibrary OrElse bLibrary Then
                                        If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblEvents(sKey).LastUpdated Then GoTo NextEvent
                                        Select Case ShouldWeLoadLibraryItem(sKey)
                                            Case LoadItemEum.Yes
                                                a.htblEvents.Remove(sKey)
                                            Case LoadItemEum.No
                                                GoTo NextEvent
                                            Case LoadItemEum.Both
                                                ' Keep key, but still add this new one
                                        End Select
                                    End If
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblEvents.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextEvent
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextEvent
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            ev.Key = sKey
                            If bLibrary Then
                                ev.IsLibrary = True
                                ev.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then ev.LastUpdated = CDate(.Item("LastUpdated").InnerText)
                            If .Item("Type") IsNot Nothing Then ev.EventType = CType([Enum].Parse(GetType(clsEvent.EventTypeEnum), .Item("Type").InnerText), clsEvent.EventTypeEnum)
                            ev.Description = .Item("Description").InnerText
                            ev.WhenStart = CType([Enum].Parse(GetType(clsEvent.WhenStartEnum), .Item("WhenStart").InnerText), clsEvent.WhenStartEnum)
                            If .Item("Repeating") IsNot Nothing Then ev.Repeating = GetBool(.Item("Repeating").InnerText)
                            If .Item("RepeatCountdown") IsNot Nothing Then ev.RepeatCountdown = GetBool(.Item("RepeatCountdown").InnerText)

                            Dim sData() As String
                            If .Item("StartDelay") IsNot Nothing Then
                                sData = .Item("StartDelay").InnerText.Split(" "c)
                                ev.StartDelay.iFrom = CInt(sData(0))
                                If sData.Length = 1 Then
                                    ev.StartDelay.iTo = CInt(sData(0))
                                Else
                                    ev.StartDelay.iTo = CInt(sData(2))
                                End If
                            End If
                            sData = .Item("Length").InnerText.Split(" "c)
                            ev.Length.iFrom = CInt(sData(0))
                            If sData.Length = 1 Then
                                ev.Length.iTo = CInt(sData(0))
                            Else
                                ev.Length.iTo = CInt(sData(2))
                            End If

                            For Each nodCtrl As XmlElement In nodEvent.GetElementsByTagName("Control")
                                Dim ctrl As New EventOrWalkControl
                                sData = nodCtrl.InnerText.Split(" "c)
                                ctrl.eControl = CType([Enum].Parse(GetType(EventOrWalkControl.ControlEnum), sData(0)), EventOrWalkControl.ControlEnum)
                                ctrl.eCompleteOrNot = CType([Enum].Parse(GetType(EventOrWalkControl.CompleteOrNotEnum), sData(1)), EventOrWalkControl.CompleteOrNotEnum)
                                ctrl.sTaskKey = sData(2)
                                ReDim Preserve ev.EventControls(ev.EventControls.Length)
                                ev.EventControls(ev.EventControls.Length - 1) = ctrl
                            Next

                            For Each nodSubEvent As XmlElement In nodEvent.GetElementsByTagName("SubEvent")
                                Dim se As New clsEvent.SubEvent(ev.Key)
                                sData = nodSubEvent.Item("When").InnerText.Split(" "c)

                                se.ftTurns.iFrom = CInt(sData(0))
                                If sData.Length = 4 Then
                                    se.ftTurns.iTo = CInt(sData(2))
                                    se.eWhen = CType([Enum].Parse(GetType(clsEvent.SubEvent.WhenEnum), sData(3).ToString), clsEvent.SubEvent.WhenEnum)
                                Else
                                    se.ftTurns.iTo = CInt(sData(0))
                                    se.eWhen = CType([Enum].Parse(GetType(clsEvent.SubEvent.WhenEnum), sData(1).ToString), clsEvent.SubEvent.WhenEnum)
                                End If

                                If nodSubEvent.Item("Action") IsNot Nothing Then
                                    If nodSubEvent.Item("Action").InnerXml.Contains("<Description>") Then
                                        se.oDescription = LoadDescription(nodSubEvent, "Action")
                                    Else
                                        sData = nodSubEvent.Item("Action").InnerText.Split(" "c)
                                        se.eWhat = CType([Enum].Parse(GetType(clsEvent.SubEvent.WhatEnum), sData(0).ToString), clsEvent.SubEvent.WhatEnum)
                                        se.sKey = sData(1)
                                    End If
                                End If

                                If nodSubEvent.Item("Measure") IsNot Nothing Then se.eMeasure = EnumParseSubEventMeasure(nodSubEvent("Measure").InnerText)
                                If nodSubEvent.Item("What") IsNot Nothing Then se.eWhat = EnumParseSubEventWhat(nodSubEvent("What").InnerText)
                                If nodSubEvent.Item("OnlyApplyAt") IsNot Nothing Then se.sKey = nodSubEvent.Item("OnlyApplyAt").InnerText

                                ReDim Preserve ev.SubEvents(ev.SubEvents.Length)
                                ev.SubEvents(ev.SubEvents.Length - 1) = se
                            Next

                        End With

                        a.htblEvents.Add(ev, ev.Key)
NextEvent:
                    Next



                    For Each nodChar As XmlElement In .SelectNodes("/Adventure/Character")
                        Dim chr As New clsCharacter
                        With nodChar
                            Dim sKey As String = .Item("Key").InnerText
                            If Not .Item("Library") Is Nothing Then chr.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblCharacters.ContainsKey(sKey) Then
                                If a.htblCharacters.ContainsKey(sKey) Then
                                    If chr.IsLibrary OrElse bLibrary Then
                                        If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblCharacters(sKey).LastUpdated Then GoTo NextChar
                                        Select Case ShouldWeLoadLibraryItem(sKey)
                                            Case LoadItemEum.Yes
                                                a.htblCharacters.Remove(sKey)
                                            Case LoadItemEum.No
                                                GoTo NextChar
                                            Case LoadItemEum.Both
                                                ' Keep key, but still add this new one
                                        End Select
                                    End If
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblCharacters.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextChar
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextChar
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            chr.Key = sKey
                            If bLibrary Then
                                chr.IsLibrary = True
                                chr.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then chr.LastUpdated = GetDate(.Item("LastUpdated").InnerText)
                            If Not .Item("Name") Is Nothing Then chr.ProperName = .Item("Name").InnerText
                            If Not .Item("Article") Is Nothing Then chr.Article = .Item("Article").InnerText
                            If Not .Item("Prefix") Is Nothing Then chr.Prefix = .Item("Prefix").InnerText
                            If .Item("Perspective") IsNot Nothing Then chr.Perspective = CType([Enum].Parse(GetType(PerspectiveEnum), .Item("Perspective").InnerText), PerspectiveEnum) Else chr.Perspective = PerspectiveEnum.ThirdPerson
                            If dFileVersion < 5.00002 Then chr.Perspective = PerspectiveEnum.SecondPerson
                            For Each nodName As XmlElement In .GetElementsByTagName("Descriptor")
                                If nodName.InnerText <> "" Then chr.arlDescriptors.Add(nodName.InnerText)
                            Next
                            chr.Description = LoadDescription(nodChar, "Description")
                            For Each nodProp As XmlElement In .SelectNodes("Property")
                                Dim prop As New clsProperty
                                Dim sPropKey As String = nodProp.Item("Key").InnerText
                                If Adventure.htblAllProperties.ContainsKey(sPropKey) Then
                                    prop = Adventure.htblAllProperties(sPropKey).Copy
                                    If prop.Type = clsProperty.PropertyTypeEnum.Text Then
                                        prop.StringData = LoadDescription(nodProp, "Value")
                                    ElseIf prop.Type <> clsProperty.PropertyTypeEnum.SelectionOnly Then
                                        prop.Value = nodProp.Item("Value").InnerText
                                    End If
                                    prop.Selected = True
                                    chr.AddProperty(prop)
                                Else
                                    ErrMsg("Error loading character " & chr.Name & ": Property " & sPropKey & " not found.")
                                End If
                            Next

                            For Each nodWalk As XmlElement In .GetElementsByTagName("Walk")
                                Dim walk As New clsWalk
                                walk.sKey = sKey
                                walk.Description = nodWalk.Item("Description").InnerText
                                walk.Loops = GetBool(nodWalk.Item("Loops").InnerText)
                                walk.StartActive = GetBool(nodWalk.Item("StartActive").InnerText)
                                For Each nodStep As XmlElement In nodWalk.GetElementsByTagName("Step")
                                    Dim [step] As New clsWalk.clsStep
                                    Dim sData() As String = nodStep.InnerText.Split(" "c)
                                    [step].sLocation = sData(0)
                                    [step].ftTurns.iFrom = CInt(sData(1))
                                    If sData.Length = 2 Then
                                        [step].ftTurns.iTo = CInt(sData(1))
                                    Else
                                        [step].ftTurns.iTo = CInt(sData(3))
                                    End If
                                    If dFileVersion < 5.000029 Then
                                        If [step].sLocation = "%Player%" Then
                                            If [step].ftTurns.iFrom < 1 Then [step].ftTurns.iFrom = 1
                                            If [step].ftTurns.iTo < 1 Then [step].ftTurns.iTo = 1
                                        End If
                                    End If
                                    walk.arlSteps.Add([step])
                                Next
                                For Each nodCtrl As XmlElement In nodWalk.GetElementsByTagName("Control")
                                    Dim ctrl As New EventOrWalkControl
                                    Dim sData() As String = nodCtrl.InnerText.Split(" "c)
                                    ctrl.eControl = CType([Enum].Parse(GetType(EventOrWalkControl.ControlEnum), sData(0)), EventOrWalkControl.ControlEnum)
                                    ctrl.eCompleteOrNot = CType([Enum].Parse(GetType(EventOrWalkControl.CompleteOrNotEnum), sData(1)), EventOrWalkControl.CompleteOrNotEnum)
                                    ctrl.sTaskKey = sData(2)
                                    ReDim Preserve walk.WalkControls(walk.WalkControls.Length)
                                    walk.WalkControls(walk.WalkControls.Length - 1) = ctrl
                                Next
                                For Each nodSubWalk As XmlElement In nodWalk.GetElementsByTagName("Activity")
                                    Dim sw As New clsWalk.SubWalk
                                    Dim sData() As String = nodSubWalk.Item("When").InnerText.Split(" "c)
                                    If sData(0) = clsWalk.SubWalk.WhenEnum.ComesAcross.ToString Then
                                        sw.eWhen = clsWalk.SubWalk.WhenEnum.ComesAcross
                                        sw.sKey = sData(1)
                                    Else
                                        sw.ftTurns.iFrom = CInt(sData(0))
                                        If sData.Length = 4 Then
                                            sw.ftTurns.iTo = CInt(sData(2))
                                            sw.eWhen = CType([Enum].Parse(GetType(clsWalk.SubWalk.WhenEnum), sData(3).ToString), clsWalk.SubWalk.WhenEnum)
                                        Else
                                            sw.ftTurns.iTo = CInt(sData(0))
                                            sw.eWhen = CType([Enum].Parse(GetType(clsWalk.SubWalk.WhenEnum), sData(1).ToString), clsWalk.SubWalk.WhenEnum)
                                        End If
                                    End If

                                    If nodSubWalk.Item("Action") IsNot Nothing Then
                                        If nodSubWalk.Item("Action").InnerXml.Contains("<Description>") Then
                                            sw.oDescription = LoadDescription(nodSubWalk, "Action")
                                        Else
                                            sData = nodSubWalk.Item("Action").InnerText.Split(" "c)
                                            sw.eWhat = CType([Enum].Parse(GetType(clsWalk.SubWalk.WhatEnum), sData(0).ToString), clsWalk.SubWalk.WhatEnum)
                                            sw.sKey2 = sData(1)
                                        End If
                                    End If

                                    If nodSubWalk.Item("OnlyApplyAt") IsNot Nothing Then sw.sKey3 = nodSubWalk.Item("OnlyApplyAt").InnerText

                                    ReDim Preserve walk.SubWalks(walk.SubWalks.Length)
                                    walk.SubWalks(walk.SubWalks.Length - 1) = sw
                                Next

                                chr.arlWalks.Add(walk)
                            Next

                            For Each nodTopic As XmlElement In .GetElementsByTagName("Topic")
                                Dim topic As New clsTopic
                                topic.Key = nodTopic.Item("Key").InnerText
                                If nodTopic.Item("ParentKey") IsNot Nothing Then topic.ParentKey = nodTopic.Item("ParentKey").InnerText
                                topic.Summary = nodTopic.Item("Summary").InnerText
                                ' Simplify Runner so it only has to deal with multiple, or specific refs                                
                                topic.Keywords = FixInitialRefs(nodTopic.Item("Keywords").InnerText)

                                topic.oConversation = LoadDescription(nodTopic, "Description")
                                If nodTopic.Item("IsAsk") IsNot Nothing Then topic.bAsk = GetBool(nodTopic.Item("IsAsk").InnerText)
                                If nodTopic.Item("IsCommand") IsNot Nothing Then topic.bCommand = GetBool(nodTopic.Item("IsCommand").InnerText)
                                If nodTopic.Item("IsFarewell") IsNot Nothing Then topic.bFarewell = GetBool(nodTopic.Item("IsFarewell").InnerText)
                                If nodTopic.Item("IsIntro") IsNot Nothing Then topic.bIntroduction = GetBool(nodTopic.Item("IsIntro").InnerText)
                                If nodTopic.Item("IsTell") IsNot Nothing Then topic.bTell = GetBool(nodTopic.Item("IsTell").InnerText)
                                If nodTopic.Item("StayInNode") IsNot Nothing Then topic.bStayInNode = GetBool(nodTopic.Item("StayInNode").InnerText)
                                topic.arlRestrictions = LoadRestrictions(nodTopic)
                                topic.arlActions = LoadActions(nodTopic)
                                chr.htblTopics.Add(topic)
                            Next

                            chr.Location = chr.Location ' Assigns the location object from the character properties
                            chr.CharacterType = EnumParseCharacterType(.Item("Type").InnerText)
                        End With
                        If eLoadWhat = LoadWhatEnum.All Then
                            a.htblCharacters.Add(chr, chr.Key)
                        Else
                            ' Only add the Player character if we don't already have one
                            If chr.CharacterType = clsCharacter.CharacterTypeEnum.NonPlayer OrElse a.Player Is Nothing Then a.htblCharacters.Add(chr, chr.Key)
                        End If
                        For Each sCharFunction As String In New String() {"CharacterName", "DisplayCharacter", "ListHeld", "ListExits", "ListWorn", "LocationOf", "ParentOf", "ProperName"}
                            chr.SearchAndReplace("%" & sCharFunction & "%", "%" & sCharFunction & "[" & chr.Key & "]%")
                        Next
NextChar:
                    Next
                    Debug.WriteLine("End Chars: " & Now)

                End If

                If eLoadWhat = LoadWhatEnum.All OrElse eLoadWhat = LoadWhatEnum.Properies Then
                    For Each nodVar As XmlElement In .SelectNodes("/Adventure/Variable")
                        Dim var As New clsVariable
                        With nodVar
                            Dim sKey As String = .Item("Key").InnerText
                            If Not .Item("Library") Is Nothing Then var.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblVariables.ContainsKey(sKey) Then
                                If var.IsLibrary OrElse bLibrary Then
                                    If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblVariables(sKey).LastUpdated Then GoTo NextVar
                                    Select Case ShouldWeLoadLibraryItem(sKey)
                                        Case LoadItemEum.Yes
                                            a.htblVariables.Remove(sKey)
                                        Case LoadItemEum.No
                                            GoTo NextVar
                                        Case LoadItemEum.Both
                                            ' Keep key, but still add this new one
                                    End Select
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblVariables.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextVar
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextVar
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            var.Key = sKey
                            If bLibrary Then
                                var.IsLibrary = True
                                var.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then var.LastUpdated = GetDate(.Item("LastUpdated").InnerText)
                            var.Name = .Item("Name").InnerText
                            var.Type = EnumParseVariableType(.Item("Type").InnerText)
                            Dim sValue As String = .Item("InitialValue").InnerText
                            If Not .Item("ArrayLength") Is Nothing Then var.Length = CInt(Val(.Item("ArrayLength").InnerText))
                            If var.Type = clsVariable.VariableTypeEnum.Text OrElse (var.Length > 1 AndAlso sValue.Contains(",")) Then
                                If var.Type = clsVariable.VariableTypeEnum.Numeric Then
                                    Dim iValues() As String = Split(sValue, ",")
                                    For iIndex As Integer = 1 To iValues.Length
                                        var.IntValue(iIndex) = SafeInt(iValues(iIndex - 1))
                                    Next
                                    var.StringValue = sValue
                                Else
                                    For iIndex As Integer = 1 To var.Length
                                        var.StringValue(iIndex) = sValue
                                    Next
                                End If

                            Else
                                For i As Integer = 1 To var.Length
                                    var.IntValue(i) = CInt(Val(sValue))
                                Next i
                            End If
                        End With
                        a.htblVariables.Add(var, var.Key)
NextVar:
                    Next
                End If

                If eLoadWhat = LoadWhatEnum.All OrElse eLoadWhat = LoadWhatEnum.AllExceptProperties Then
                    For Each nodGroup As XmlElement In .SelectNodes("/Adventure/Group")
                        Dim grp As New clsGroup
                        With nodGroup
                            Dim sKey As String = .Item("Key").InnerText
                            If Not .Item("Library") Is Nothing Then grp.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblGroups.ContainsKey(sKey) Then
                                If grp.IsLibrary OrElse bLibrary Then
                                    If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblGroups(sKey).LastUpdated Then GoTo NextGroup
                                    Select Case ShouldWeLoadLibraryItem(sKey)
                                        Case LoadItemEum.Yes
                                            a.htblGroups.Remove(sKey)
                                        Case LoadItemEum.No
                                            GoTo NextGroup
                                        Case LoadItemEum.Both
                                            ' Keep key, but still add this new one
                                    End Select
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblGroups.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextGroup
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextGroup
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            grp.Key = sKey
                            If bLibrary Then
                                grp.IsLibrary = True
                                grp.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then grp.LastUpdated = GetDate(.Item("LastUpdated").InnerText)
                            grp.Name = .Item("Name").InnerText
                            grp.GroupType = EnumParseGroupType(.Item("Type").InnerText)
                            For Each nodMember As XmlElement In .GetElementsByTagName("Member")
                                grp.arlMembers.Add(nodMember.InnerText)
                            Next
                            For Each nodProp As XmlElement In .GetElementsByTagName("Property")
                                Dim prop As New clsProperty
                                Dim sPropKey As String = nodProp.Item("Key").InnerText
                                If Adventure.htblAllProperties.ContainsKey(sPropKey) Then
                                    prop = Adventure.htblAllProperties(sPropKey).Copy
                                    If nodProp.Item("Value") IsNot Nothing Then
                                        If prop.Type = clsProperty.PropertyTypeEnum.Text Then
                                            prop.StringData = LoadDescription(nodProp, "Value")
                                        ElseIf prop.Type <> clsProperty.PropertyTypeEnum.SelectionOnly Then
                                            prop.Value = nodProp.Item("Value").InnerText
                                        End If
                                    End If
                                    prop.Selected = True
                                    If prop.PropertyOf = grp.GroupType Then
                                        grp.htblProperties.Add(prop)
                                    End If
                                Else
                                    ErrMsg("Error loading group " & grp.Name & ": Property " & sPropKey & " not found.")
                                End If
                            Next
                        End With
                        a.htblGroups.Add(grp, grp.Key)

                        For Each sMember As String In grp.arlMembers
                            Dim itm As clsItemWithProperties = CType(Adventure.GetItemFromKey(sMember), clsItemWithProperties)
                            If itm IsNot Nothing Then itm.ResetInherited() ' In case we've accessed properties, and built inherited before the group existed
                        Next
NextGroup:
                    Next


                    For Each nodALR As XmlElement In .SelectNodes("/Adventure/TextOverride")
                        Dim alr As New clsALR
                        With nodALR
                            Dim sKey As String = .Item("Key").InnerText
                            If .Item("Library") IsNot Nothing Then alr.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblALRs.ContainsKey(sKey) Then
                                If alr.IsLibrary OrElse bLibrary Then
                                    If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblALRs(sKey).LastUpdated Then GoTo NextALR
                                    Select Case ShouldWeLoadLibraryItem(sKey)
                                        Case LoadItemEum.Yes
                                            a.htblALRs.Remove(sKey)
                                        Case LoadItemEum.No
                                            GoTo NextALR
                                        Case LoadItemEum.Both
                                            ' Keep key, but still add this new one
                                    End Select
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblALRs.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextALR
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextALR
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            alr.Key = sKey
                            If bLibrary Then
                                alr.IsLibrary = True
                                alr.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then alr.LastUpdated = CDate(.Item("LastUpdated").InnerText)
                            alr.OldText = CStr(.Item("OldText").InnerText)
                            alr.NewText = LoadDescription(nodALR, "NewText")
                        End With
                        a.htblALRs.Add(alr, alr.Key)
NextALR:
                    Next

                    For Each nodHint As XmlElement In .SelectNodes("/Adventure/Hint")
                        Dim hint As New clsHint
                        With nodHint
                            Dim sKey As String = .Item("Key").InnerText
                            If .Item("Library") IsNot Nothing Then hint.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblHints.ContainsKey(sKey) Then
                                If hint.IsLibrary OrElse bLibrary Then
                                    If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblHints(sKey).LastUpdated Then GoTo NextHint
                                    Select Case ShouldWeLoadLibraryItem(sKey)
                                        Case LoadItemEum.Yes
                                            a.htblHints.Remove(sKey)
                                        Case LoadItemEum.No
                                            GoTo NextHint
                                        Case LoadItemEum.Both
                                            ' Keep key, but still add this new one
                                    End Select
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblHints.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextHint
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextHint
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            hint.Key = sKey
                            If bLibrary Then
                                hint.IsLibrary = True
                                hint.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then hint.LastUpdated = CDate(.Item("LastUpdated").InnerText)
                            hint.Question = CStr(.Item("Question").InnerText)
                            hint.SubtleHint = LoadDescription(nodHint, "Subtle")
                            hint.SledgeHammerHint = LoadDescription(nodHint, "Sledgehammer")
                            hint.arlRestrictions = LoadRestrictions(nodHint)
                        End With
                        a.htblHints.Add(hint, hint.Key)
NextHint:
                    Next

                    For Each nodSynonym As XmlElement In .SelectNodes("/Adventure/Synonym")
                        Dim synonym As New clsSynonym
                        With nodSynonym
                            Dim sKey As String = .Item("Key").InnerText
                            If .Item("Library") IsNot Nothing Then synonym.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblSynonyms.ContainsKey(sKey) Then
                                If synonym.IsLibrary OrElse bLibrary Then
                                    If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblSynonyms(sKey).LastUpdated Then GoTo NextSynonym
                                    Select Case ShouldWeLoadLibraryItem(sKey)
                                        Case LoadItemEum.Yes
                                            a.htblSynonyms.Remove(sKey)
                                        Case LoadItemEum.No
                                            GoTo NextSynonym
                                        Case LoadItemEum.Both
                                            ' Keep key, but still add this new one
                                    End Select
                                End If

                                If bAddDuplicateKeys Then
                                    While a.htblSynonyms.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextSynonym
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextSynonym
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            synonym.Key = sKey
                            If bLibrary Then
                                synonym.IsLibrary = True
                                synonym.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then synonym.LastUpdated = CDate(.Item("LastUpdated").InnerText)

                            For Each nodFrom As XmlElement In .GetElementsByTagName("From")
                                synonym.ChangeFrom.Add(nodFrom.InnerText)
                            Next
                            synonym.ChangeTo = .Item("To").InnerText
                        End With
                        a.htblSynonyms.Add(synonym)
NextSynonym:
                    Next


                    For Each nodUDF As XmlElement In .SelectNodes("/Adventure/Function")
                        Dim udf As New clsUserFunction
                        With nodUDF
                            Dim sKey As String = .Item("Key").InnerText
                            If .Item("Library") IsNot Nothing Then udf.IsLibrary = GetBool(.Item("Library").InnerText)
                            If a.htblUDFs.ContainsKey(sKey) Then
                                If udf.IsLibrary OrElse bLibrary Then
                                    If .Item("LastUpdated") Is Nothing OrElse CDate(.Item("LastUpdated").InnerText) <= a.htblUDFs(sKey).LastUpdated Then GoTo NextUDF
                                    Select Case ShouldWeLoadLibraryItem(sKey)
                                        Case LoadItemEum.Yes
                                            a.htblUDFs.Remove(sKey)
                                        Case LoadItemEum.No
                                            GoTo NextUDF
                                        Case LoadItemEum.Both
                                            ' Keep key, but still add this new one
                                    End Select
                                End If
                                If bAddDuplicateKeys Then
                                    While a.htblUDFs.ContainsKey(sKey)
                                        sKey = IncrementKey(sKey)
                                    End While
                                Else
                                    GoTo NextUDF
                                End If
                            ElseIf bLibrary AndAlso ShouldWeLoadLibraryItem(sKey) = LoadItemEum.No Then
                                GoTo NextUDF
                            End If
                            If a.listExcludedItems.Contains(sKey) Then a.listExcludedItems.Remove(sKey)
                            udf.Key = sKey
                            If bLibrary Then
                                udf.IsLibrary = True
                                udf.LoadedFromLibrary = True
                            End If
                            If Not .Item("LastUpdated") Is Nothing Then udf.LastUpdated = CDate(.Item("LastUpdated").InnerText)

                            udf.Name = CStr(.Item("Name").InnerText)
                            udf.Output = LoadDescription(nodUDF, "Output")
                            For Each nodArgument As XmlElement In .GetElementsByTagName("Argument")
                                With nodArgument
                                    Dim arg As New clsUserFunction.Argument
                                    arg.Name = .Item("Name").InnerText
                                    arg.Type = CType([Enum].Parse(GetType(clsUserFunction.ArgumentType), .Item("Type").InnerText), clsUserFunction.ArgumentType)
                                    udf.Arguments.Add(arg)
                                End With
                            Next

                        End With
                        a.htblUDFs.Add(udf, udf.Key)
NextUDF:
                    Next

                    If Not bLibrary Then
                        For Each nodExclude As XmlElement In .SelectNodes("/Adventure/Exclude")
                            a.listExcludedItems.Add(nodExclude.InnerText)
                        Next
                    End If

                    If Not bLibrary Then
                        If .Item("Map") IsNot Nothing Then
                            Adventure.Map.Pages.Clear()
                            For Each nodPage As XmlElement In .SelectNodes("/Adventure/Map/Page")
                                With nodPage
                                    Dim sPageKey As String = .Item("Key").InnerText
                                    If IsNumeric(sPageKey) Then
                                        Dim page As New MapPage(SafeInt(sPageKey))
                                        If .Item("Label") IsNot Nothing Then page.Label = .Item("Label").InnerText
                                        If .Item("Selected") IsNot Nothing AndAlso GetBool(.Item("Selected").InnerText) Then Adventure.Map.SelectedPage = sPageKey
                                        For Each nodNode As XmlElement In .GetElementsByTagName("Node")
                                            With nodNode
                                                Dim node As New MapNode
                                                If .Item("Key") IsNot Nothing Then node.Key = .Item("Key").InnerText
                                                Dim loc As clsLocation = Adventure.htblLocations(node.Key)
                                                If loc IsNot Nothing Then
                                                    node.Text = loc.ShortDescriptionSafe ' StripCarats(ReplaceALRs(loc.ShortDescription.ToString))
                                                    If .Item("X") IsNot Nothing Then node.Location.X = SafeInt(.Item("X").InnerText)
                                                    If .Item("Y") IsNot Nothing Then node.Location.Y = SafeInt(.Item("Y").InnerText)
                                                    If .Item("Z") IsNot Nothing Then node.Location.Z = SafeInt(.Item("Z").InnerText)
                                                    If .Item("Height") IsNot Nothing Then node.Height = SafeInt(.Item("Height").InnerText) Else node.Height = 4
                                                    If .Item("Width") IsNot Nothing Then node.Width = SafeInt(.Item("Width").InnerText) Else node.Width = 6

                                                    If loc.arlDirections(DirectionsEnum.Up).LocationKey IsNot Nothing Then node.bHasUp = True
                                                    If loc.arlDirections(DirectionsEnum.Down).LocationKey IsNot Nothing Then node.bHasDown = True
                                                    If loc.arlDirections(DirectionsEnum.In).LocationKey IsNot Nothing Then node.bHasIn = True
                                                    If loc.arlDirections(DirectionsEnum.Out).LocationKey IsNot Nothing Then node.bHasOut = True

                                                    For Each nodLink As XmlElement In .GetElementsByTagName("Link")
                                                        With nodLink
                                                            Dim link As New MapLink
                                                            link.sSource = node.Key
                                                            If .Item("SourceAnchor") IsNot Nothing Then link.eSourceLinkPoint = CType([Enum].Parse(GetType(DirectionsEnum), .Item("SourceAnchor").InnerText), DirectionsEnum)
                                                            link.sDestination = loc.arlDirections(link.eSourceLinkPoint).LocationKey
                                                            If Adventure.Map.DottedLink(loc.arlDirections(link.eSourceLinkPoint)) Then
                                                                link.Style = DashStyles.Dot
                                                            Else
                                                                link.Style = DashStyles.Solid
                                                            End If
                                                            If .Item("DestinationAnchor") IsNot Nothing Then link.eDestinationLinkPoint = CType([Enum].Parse(GetType(DirectionsEnum), .Item("DestinationAnchor").InnerText), DirectionsEnum)
                                                            Dim sDest As String = loc.arlDirections(link.eSourceLinkPoint).LocationKey
                                                            If sDest IsNot Nothing AndAlso Adventure.htblLocations.ContainsKey(sDest) Then
                                                                Dim locDest As clsLocation = Adventure.htblLocations(sDest)
                                                                If locDest IsNot Nothing Then
                                                                    If locDest.arlDirections(link.eDestinationLinkPoint).LocationKey = loc.Key Then
                                                                        link.Duplex = True
                                                                        If Adventure.Map.DottedLink(locDest.arlDirections(link.eDestinationLinkPoint)) Then link.Style = DashStyles.Dot
                                                                    End If
                                                                End If
                                                            End If
                                                            For Each nodAnchor As XmlElement In .GetElementsByTagName("Anchor")
                                                                Dim p As New Point3D
                                                                With nodAnchor
                                                                    If .Item("X") IsNot Nothing Then p.X = SafeInt(.Item("X").InnerText)
                                                                    If .Item("Y") IsNot Nothing Then p.Y = SafeInt(.Item("Y").InnerText)
                                                                    If .Item("Z") IsNot Nothing Then p.Z = SafeInt(.Item("Z").InnerText)
                                                                End With
                                                                ReDim Preserve link.OrigMidPoints(link.OrigMidPoints.Length)
                                                                link.OrigMidPoints(link.OrigMidPoints.Length - 1) = p
                                                                Dim ar As New Anchor
                                                                ar.Visible = True
                                                                ar.Parent = link
                                                                link.Anchors.Add(ar)
                                                            Next
                                                            node.Links.Add(link.eSourceLinkPoint, link)
                                                        End With
                                                    Next

                                                    node.Page = page.iKey
                                                    page.AddNode(node)
                                                End If
                                            End With
                                        Next
                                        Adventure.Map.Pages.Add(page.iKey, page)
                                        Adventure.Map.Pages(page.iKey).SortNodes()
                                    End If
                                End With
                            Next
                        End If
                    End If

                    ' Now fix any remapped keys
                    ' This must only remap our newly imported tasks, not all the original ones!
                    '
                    For Each sOldKey As String In htblDuplicateKeyMapping.Keys
                        For Each sTask As String In arlNewTasks
                            Dim tas As clsTask = a.htblTasks(sTask)
                            If tas.GeneralKey = sOldKey Then tas.GeneralKey = htblDuplicateKeyMapping(sOldKey)
                            For Each act As clsAction In tas.arlActions
                                If act.sKey1 = sOldKey Then act.sKey1 = htblDuplicateKeyMapping(sOldKey)
                                If act.sKey2 = sOldKey Then act.sKey2 = htblDuplicateKeyMapping(sOldKey)
                            Next
                        Next
                    Next

                    Adventure.BlorbMappings.Clear()
                    If .Item("FileMappings") IsNot Nothing Then
                        For Each nodMapping As XmlElement In .SelectNodes("/Adventure/FileMappings/Mapping")
                            With nodMapping
                                Dim iResource As Integer = SafeInt(.Item("Resource").InnerText)
                                Dim sFile As String = .Item("File").InnerText
                                Adventure.BlorbMappings.Add(sFile, iResource)
                            End With
                        Next
                    End If

                End If

            End With

            ' Correct any old style functions
            ' Player.Held.Weight > Player.Held.Weight.Sum
            If Not bLibrary AndAlso dFileVersion < 5.0000311 Then
                For Each sd As SingleDescription In Adventure.AllDescriptions
                    For Each p As clsProperty In Adventure.htblAllProperties.Values
                        If p.Type = clsProperty.PropertyTypeEnum.Integer OrElse p.Type = clsProperty.PropertyTypeEnum.ValueList Then
                            If sd.Description.Contains("." & p.Key) Then
                                sd.Description = sd.Description.Replace("." & p.Key, "." & p.Key & ".Sum")
                            End If
                        End If
                    Next
                Next
            End If


            If eLoadWhat = LoadWhatEnum.All AndAlso iCorrectedTasks > 0 Then
                Glue.ShowInfo(iCorrectedTasks & " tasks have been updated.", "Adventure Upgrade")
            End If

            With Adventure
                If .Map.Pages.Count = 1 AndAlso .Map.Pages.ContainsKey(0) AndAlso .Map.Pages(0).Nodes.Count = 0 Then
                    .Map = New clsMap
                    .Map.RecalculateLayout()
                End If
            End With
            dFileVersion = dVersion ' Set back to current version so copy/paste etc has correct versions

            Return True

        Catch exLAE As LoadAbortException
            ' Ignore
            Return False
        Catch exXML As XmlException
            If bLibrary Then
                ErrMsg("The file '" & sFilename & "' you are trying to load is not a valid ADRIFT Module.", exXML)
            Else
                ErrMsg("Error loading Adventure", exXML)
            End If
        Catch ex As Exception
            If ex.Message.Contains("Root element is missing") Then
                ErrMsg("The file you are trying to load is not a valid ADRIFT v5.0 file.")
            Else
                ErrMsg("Error loading Adventure", ex)
            End If
            Return False
        End Try

    End Function


    Private Function LocRestriction(ByVal sLocKey As String, ByVal bMust As Boolean) As clsRestriction

        Dim r As New clsRestriction
        r.eType = clsRestriction.RestrictionTypeEnum.Character
        r.sKey1 = "%Player%"
        If bMust Then
            r.eMust = clsRestriction.MustEnum.Must
        Else
            r.eMust = clsRestriction.MustEnum.MustNot
        End If
        If Adventure.htblLocations.ContainsKey(sLocKey) Then
            r.eCharacter = clsRestriction.CharacterEnum.BeAtLocation
        ElseIf Adventure.htblGroups.ContainsKey(sLocKey) Then
            r.eCharacter = clsRestriction.CharacterEnum.BeWithinLocationGroup
        End If
        r.sKey2 = sLocKey
        Return r

    End Function


    ' Converts v4 functions to v5
    Private Function ConvText(ByVal s400Text As String) As String
        Dim s500Text As String = s400Text.Replace("%theobject%", "%TheObject[%object%]%").Replace("%character%", "%CharacterName[%character%]%").Replace("%room%", "%LocationName[%LocationOf[Player]%]%").Replace("%obstatus%", "%LCase[%PropertyValue[%object%,OpenStatus]%]%")
        While s500Text.Contains("%t_")
            Dim iStart As Integer = s500Text.IndexOf("%t_")
            Dim iEnd As Integer = s500Text.IndexOf("%", iStart + 1) + 1
            s500Text = s500Text.Replace(s500Text.Substring(iStart, iEnd - iStart), "%NumberAsText[%" & s500Text.Substring(iStart + 3, iEnd - iStart - 3) & "]%")
        End While
        While s500Text.Contains("%in_")
            Dim iStart As Integer = s500Text.IndexOf("%in_")
            Dim iEnd As Integer = s500Text.IndexOf("%", iStart + 1) + 1
            ' convert object name to key
            Dim sKey As String = ""
            For Each ob As clsObject In Adventure.htblObjects.Values
                If s500Text.Substring(iStart + 4, iEnd - iStart - 5) = ob.arlNames(0) Then
                    sKey = ob.Key
                    Exit For
                End If
            Next
            s500Text = s500Text.Replace(s500Text.Substring(iStart, iEnd - iStart), "%ListObjectsIn[" & sKey & "]%")
        End While
        While s500Text.Contains("%on_")
            Dim iStart As Integer = s500Text.IndexOf("%on_")
            Dim iEnd As Integer = s500Text.IndexOf("%", iStart + 1) + 1
            ' convert object name to key
            Dim sKey As String = ""
            For Each ob As clsObject In Adventure.htblObjects.Values
                If s500Text.Substring(iStart + 4, iEnd - iStart - 5) = ob.arlNames(0) Then
                    sKey = ob.Key
                    Exit For
                End If
            Next
            s500Text = s500Text.Replace(s500Text.Substring(iStart, iEnd - iStart), "%ListObjectsOn[" & sKey & "]%")
        End While
        While s500Text.Contains("%onin_")
            Dim iStart As Integer = s500Text.IndexOf("%onin_")
            Dim iEnd As Integer = s500Text.IndexOf("%", iStart + 1) + 1
            ' convert object name to key
            Dim sKey As String = ""
            For Each ob As clsObject In Adventure.htblObjects.Values
                If s500Text.Substring(iStart + 6, iEnd - iStart - 7) = ob.arlNames(0) Then
                    sKey = ob.Key
                    Exit For
                End If
            Next
            s500Text = s500Text.Replace(s500Text.Substring(iStart, iEnd - iStart), "%ListObjectsOnAndIn[" & sKey & "]%")
        End While
        ' %status_
        ' %state_
        Return s500Text
    End Function

    Friend Function CopyStream(input As Stream, output As Stream) As Boolean
        Try
            Dim iBlock As Integer = 1024
            Dim iBytes As Integer
            Dim buffer1 As Byte() = New Byte(iBlock - 1) {}
            iBytes = input.Read(buffer1, 0, iBlock)
            Do While (iBytes > 0)
                output.Write(buffer1, 0, iBytes)
                iBytes = input.Read(buffer1, 0, iBlock)
            Loop
            output.Flush()
            Return True
        Catch ex As SharpZipBaseException
            ErrMsg("CopyStream error", ex)
            Return False
        End Try
    End Function

    Friend Sub CreateMandatoryProperties()
        For Each sKey As String In New String() {OBJECTARTICLE, OBJECTPREFIX, OBJECTNOUN}
            If Not Adventure.htblObjectProperties.ContainsKey(sKey) Then
                Dim prop As New clsProperty
                With prop
                    .Key = sKey
                    Select Case .Key
                        Case OBJECTARTICLE
                            .Description = "Object Article"
                        Case OBJECTPREFIX
                            .Description = "Object Prefix"
                        Case OBJECTNOUN
                            .Description = "Object Name"
                    End Select
                    .PropertyOf = clsProperty.PropertyOfEnum.Objects
                    .Type = clsProperty.PropertyTypeEnum.Text
                End With
                Adventure.htblAllProperties.Add(prop)
            End If
            Adventure.htblObjectProperties(sKey).GroupOnly = True
        Next

        If Not Adventure.htblLocationProperties.ContainsKey(SHORTLOCATIONDESCRIPTION) Then
            Dim prop As New clsProperty
            With prop
                .Key = SHORTLOCATIONDESCRIPTION
                .Description = "Short Location Description"
                .PropertyOf = clsProperty.PropertyOfEnum.Locations
                .Type = clsProperty.PropertyTypeEnum.Text
            End With
            Adventure.htblAllProperties.Add(prop)
        End If
        Adventure.htblLocationProperties(SHORTLOCATIONDESCRIPTION).GroupOnly = True

        If Not Adventure.htblLocationProperties.ContainsKey(LONGLOCATIONDESCRIPTION) Then
            Dim prop As New clsProperty
            With prop
                .Key = LONGLOCATIONDESCRIPTION
                .Description = "Long Location Description"
                .PropertyOf = clsProperty.PropertyOfEnum.Locations
                .Type = clsProperty.PropertyTypeEnum.Text
            End With
            Adventure.htblAllProperties.Add(prop)
        End If
        Adventure.htblLocationProperties(LONGLOCATIONDESCRIPTION).GroupOnly = True

        If Not Adventure.htblCharacterProperties.ContainsKey(CHARACTERPROPERNAME) Then
            Dim prop As New clsProperty
            With prop
                .Key = CHARACTERPROPERNAME
                .Description = "Character Proper Name"
                .PropertyOf = clsProperty.PropertyOfEnum.Characters
                .Type = clsProperty.PropertyTypeEnum.Text
            End With
            Adventure.htblAllProperties.Add(prop)
        End If
        Adventure.htblCharacterProperties(CHARACTERPROPERNAME).GroupOnly = True

        If Not Adventure.htblObjectProperties.ContainsKey("StaticOrDynamic") Then
            Dim prop As New clsProperty
            With prop
                .Key = "StaticOrDynamic"
                .Description = "Object type"
                .Mandatory = True
                .PropertyOf = clsProperty.PropertyOfEnum.Objects
                .Type = clsProperty.PropertyTypeEnum.StateList
                .arlStates.Add("Static")
                .arlStates.Add("Dynamic")
            End With
            Adventure.htblAllProperties.Add(prop)
        End If
        Adventure.htblObjectProperties("StaticOrDynamic").GroupOnly = True

        If Not Adventure.htblObjectProperties.ContainsKey("StaticLocation") Then
            Dim prop As New clsProperty
            With prop
                .Key = "StaticLocation"
                .Description = "Location of the object"
                .Mandatory = True
                .PropertyOf = clsProperty.PropertyOfEnum.Objects
                .Type = clsProperty.PropertyTypeEnum.StateList
                .arlStates.Add("Hidden")
                .arlStates.Add("Single Location")
                .arlStates.Add("Location Group")
                .arlStates.Add("Everywhere")
                .arlStates.Add("Part of Character")
                .arlStates.Add("Part of Object")
                .DependentKey = "StaticOrDynamic"
                .DependentValue = "Static"
            End With
            Adventure.htblAllProperties.Add(prop)
        End If
        Adventure.htblObjectProperties("StaticLocation").GroupOnly = True

        If Not Adventure.htblObjectProperties.ContainsKey("DynamicLocation") Then
            Dim prop As New clsProperty
            With prop
                .Key = "DynamicLocation"
                .Description = "Location of the object"
                .Mandatory = True
                .PropertyOf = clsProperty.PropertyOfEnum.Objects
                .Type = clsProperty.PropertyTypeEnum.StateList
                .arlStates.Add("Hidden")
                .arlStates.Add("Held by Character")
                .arlStates.Add("Worn by Character")
                .arlStates.Add("In Location")
                .arlStates.Add("Inside Object")
                .arlStates.Add("On Object")
                .DependentKey = "StaticOrDynamic"
                .DependentValue = "Dynamic"
            End With
            Adventure.htblAllProperties.Add(prop)
        End If
        Adventure.htblObjectProperties("DynamicLocation").GroupOnly = True

        For Each sProp As String In New String() {"AtLocation", "AtLocationGroup", "PartOfWhat", "PartOfWho", "HeldByWho", "WornByWho", "InLocation", "InsideWhat", "OnWhat"}
            If Adventure.htblObjectProperties.ContainsKey(sProp) Then Adventure.htblObjectProperties(sProp).GroupOnly = True
        Next

    End Sub

    Private Function LoadOlder(ByVal v As Double) As Boolean
        Try
            If Adventure.htblAllProperties.Count = 0 Then
                ErrMsg("You must select at least one library within Generator > File > Settings > Libraries before loading ADRIFT v" & v.ToString("#.0") & " adventures.")
                Return False
            End If

            ' Safety check...
            Dim sPropCheck As String = ""
            For Each sProperty As String In New String() {"AtLocation", "AtLocationGroup", "CharacterAtLocation", "CharacterLocation", "CharEnters", "CharExits", "Container", "DynamicLocation", "HeldByWho", "InLocation", "InsideWhat", "Lockable", "LockKey", "LockStatus", "ListDescription", "OnWhat", "Openable", "OpenStatus", "PartOfWho", "Readable", "ReadText", "ShowEnterExit", "StaticLocation", "StaticOrDynamic", "Surface", "Wearable", "WornByWho"}
                If Not Adventure.htblAllProperties.ContainsKey(sProperty) Then
                    sPropCheck &= sProperty & vbCrLf
                End If
            Next
            If sPropCheck <> "" Then
                ErrMsg("Library must contain the following properties before loading ADRIFT v" & v.ToString("#.0") & " files:" & vbCrLf & sPropCheck)
                Return False
            End If

            Dim bAdvZLib() As Byte = Nothing
            If v < 4 Then
                br.BaseStream.Position = 0
                Dim bRawData() As Byte = br.ReadBytes(CInt(br.BaseStream.Length))
                Dim sRawData As String = System.Text.Encoding.Default.GetString(bRawData)
                bAdventure = Dencode(sRawData, 0)
            Else
                br.ReadBytes(2) ' CrLf
                Dim lSize As Long = CLng(System.Text.Encoding.Default.GetString(br.ReadBytes(8)))
                bAdvZLib = br.ReadBytes(CInt(lSize - 23))
                Dim sPassword As String = System.Text.Encoding.Default.GetString(Dencode(System.Text.Encoding.Default.GetString(br.ReadBytes(12)), lSize + 1))


                ReDim bAdventure(-1)

                Dim outStream As New System.IO.MemoryStream
                Dim inStream As New System.IO.MemoryStream(bAdvZLib)
                Dim zStream As New InflaterInputStream(inStream)
                Try
                    CopyStream(zStream, outStream)
                    bAdventure = outStream.ToArray
                Finally
                    zStream.Close()
                    outStream.Close()
                    inStream.Close()
                End Try
            End If

            Dim a As clsAdventure = Adventure
            With Adventure

                Dim iStartMaxPriority As Integer = CurrentMaxPriority()
                Dim iStartLocations As Integer = .htblLocations.Count
                Dim iStartObs As Integer = .htblObjects.Count
                Dim iStartTask As Integer = .htblTasks.Count
                Dim iStartChar As Integer = .htblCharacters.Count
                Dim iStartVariable As Integer = .htblVariables.Count
                Dim bSound As Boolean = False
                Dim bGraphics As Boolean = False
                Dim sFilename As String
                Dim iFilesize As Integer
                .dictv4Media = New Dictionary(Of String, clsAdventure.v4Media)

                .TaskExecution = clsAdventure.TaskExecutionEnum.HighestPriorityPassingTask

                salWithStates.Clear()
                Dim iPos As Integer = 0
                Dim sBuffer As String = Nothing
                ' Read the introduction

                Dim sTerminator As String
                If v < 4 Then
                    sTerminator = "**"
                Else
                    sTerminator = "��"
                End If


                While iPos < bAdventure.Length - 1 And sBuffer <> sTerminator
                    sBuffer = GetLine(bAdventure, iPos)
                    If sBuffer <> sTerminator Then .Introduction.Item(0).Description &= "<br>" & sBuffer
                End While
                Dim iStartLocation As Integer = CInt(GetLine(bAdventure, iPos))
                sBuffer = Nothing
                While iPos < bAdventure.Length - 1 And sBuffer <> sTerminator
                    sBuffer = GetLine(bAdventure, iPos)
                    If sBuffer <> sTerminator Then .WinningText.Item(0).Description &= "<br>" & sBuffer
                End While
                .Title = GetLine(bAdventure, iPos)
                .Author = GetLine(bAdventure, iPos)
                If v < 3.9 Then
                    Dim sNoOb As String = GetLine(bAdventure, iPos)
                End If
                .NotUnderstood = GetLine(bAdventure, iPos)
                Dim iPer As Integer = CInt(GetLine(bAdventure, iPos)) + 1 ' Perspective
                .ShowExits = CBool(GetLine(bAdventure, iPos))
                .WaitTurns = SafeInt(GetLine(bAdventure, iPos)) ' WaitNum

                Dim iCompassPoints As Integer = 8
                Dim bBattleSystem As Boolean = False

                If v >= 3.9 Then
                    .ShowFirstRoom = CBool(GetLine(bAdventure, iPos)) ' ShowFirst
                    bBattleSystem = CBool(GetLine(bAdventure, iPos))
                    .MaxScore = CInt(GetLine(bAdventure, iPos))

                    ' Delete the default char in new file
                    .htblCharacters.Remove("Player")
                    Dim Player As New clsCharacter
                    With Player
                        .Key = "Player"
                        .CharacterType = clsCharacter.CharacterTypeEnum.Player
                        .Perspective = CType(iPer, PerspectiveEnum)
                        .ProperName = GetLine(bAdventure, iPos)
                        If .ProperName = "" OrElse .ProperName = "Anonymous" Then .ProperName = "Player"
                        Dim bPromptName As Boolean = CBool(GetLine(bAdventure, iPos))  ' PromptName       
                        .arlDescriptors.Add("myself")
                        .arlDescriptors.Add("me")

                        Dim p As New clsProperty
                        p = Adventure.htblCharacterProperties("Known").Copy
                        p.Selected = True
                        .AddProperty(p)

                        .Description = New Description(ConvText(GetLine(bAdventure, iPos))) ' Description
                        Dim iTask As Integer = CInt(GetLine(bAdventure, iPos)) ' Task
                        If iTask > 0 Then
                            GetLine(bAdventure, iPos) ' Description
                        End If
                        .Location.Position = CType(GetLine(bAdventure, iPos), clsCharacterLocation.PositionEnum)
                        Dim iOnWhat As Integer = CInt(GetLine(bAdventure, iPos)) ' OnWhat                    
                        Dim bPromptGender As Boolean = False
                        Select Case SafeInt(GetLine(bAdventure, iPos)) ' Sex
                            Case 0 ' Male
                                .Gender = clsCharacter.GenderEnum.Male
                            Case 1 ' Female
                                .Gender = clsCharacter.GenderEnum.Female
                            Case 2 ' Prompt
                                .Gender = clsCharacter.GenderEnum.Unknown
                                bPromptGender = True
                        End Select

                        If bPromptName OrElse bPromptGender Then
                            Dim tasPrompt As New clsTask
                            With tasPrompt
                                .Key = "GenTask" & (Adventure.htblTasks.Count + 1) ' Give it a unique key so it doesn't disrupt events calling tasks by index
                                .Description = "Generated task for Player prompts"
                                .TaskType = clsTask.TaskTypeEnum.System
                                .RunImmediately = True
                                .Priority = iStartMaxPriority + Adventure.htblTasks.Count + 1
                                If bPromptGender Then
                                    Dim act As New clsAction
                                    act.eItem = clsAction.ItemEnum.SetProperties
                                    act.sKey1 = Player.Key
                                    act.sKey2 = "Gender"
                                    act.sPropertyValue = "%PopUpChoice[""Please select player gender"", ""Male"", ""Female""]%"
                                    tasPrompt.arlActions.Add(act)
                                End If
                                If bPromptName Then
                                    Dim act As New clsAction
                                    act.eItem = clsAction.ItemEnum.SetProperties
                                    act.sKey1 = Player.Key
                                    act.sKey2 = "CharacterProperName"
                                    act.sPropertyValue = "%PopUpInput[""Please enter your name"", ""Anonymous""]%"
                                    tasPrompt.arlActions.Add(act)
                                End If
                            End With
                            Adventure.htblTasks.Add(tasPrompt, tasPrompt.Key)
                        End If
                        .MaxSize = CInt(GetLine(bAdventure, iPos))

                        If Adventure.htblCharacterProperties.ContainsKey("MaxBulk") Then
                            p = New clsProperty
                            p = Adventure.htblCharacterProperties("MaxBulk").Copy
                            p.Selected = True
                            Dim iSize2 As Integer = .MaxSize Mod 10
                            p.Value = (CInt((.MaxSize - iSize2) / 10) * (3 ^ iSize2)).ToString
                            .AddProperty(p)
                        End If

                        .MaxWeight = CInt(GetLine(bAdventure, iPos))

                        If Adventure.htblCharacterProperties.ContainsKey("MaxWeight") Then
                            p = New clsProperty
                            p = Adventure.htblCharacterProperties("MaxWeight").Copy
                            p.Selected = True
                            Dim iWeight2 As Integer = .MaxWeight Mod 10
                            p.Value = (CInt((.MaxWeight - iWeight2) / 10) * (3 ^ iWeight2)).ToString
                            .AddProperty(p)
                        End If

                        If iOnWhat = 0 Then
                            .Location.ExistWhere = clsCharacterLocation.ExistsWhereEnum.AtLocation
                            .Location.Key = "Location" & iStartLocation + 1
                        Else
                            .Location.ExistWhere = clsCharacterLocation.ExistsWhereEnum.OnObject
                            .Location.Key = iOnWhat.ToString ' Will adjust this below
                        End If


                        .Location = .Location

                        If bBattleSystem Then
                            GetLine(bAdventure, iPos) ' Min Stamina
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Max Stamina
                            GetLine(bAdventure, iPos) ' Min Strength
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Max Strength
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Min Accuracy
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Max Accuracy
                            GetLine(bAdventure, iPos) ' Min Defence
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Max Defence
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Min Agility
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Max Agility
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Recovery
                        End If
                    End With
                    .htblCharacters.Add(Player, Player.Key)

                    'Adventure.iCompassPoints = CType(8 + 4 * CInt(GetLine(bAdventure, iPos)), DirectionsEnum)
                    iCompassPoints = CType(8 + 4 * CInt(GetLine(bAdventure, iPos)), DirectionsEnum) ' CInt(GetLine(bAdventure, iPos))
                    Adventure.Enabled(clsAdventure.EnabledOptionEnum.Debugger) = Not CBool(GetLine(bAdventure, iPos)) '?
                    Adventure.Enabled(clsAdventure.EnabledOptionEnum.Score) = Not CBool(GetLine(bAdventure, iPos)) '?
                    Adventure.Enabled(clsAdventure.EnabledOptionEnum.Map) = Not CBool(GetLine(bAdventure, iPos))
                    Adventure.Enabled(clsAdventure.EnabledOptionEnum.AutoComplete) = Not CBool(GetLine(bAdventure, iPos)) '?
                    Adventure.Enabled(clsAdventure.EnabledOptionEnum.ControlPanel) = Not CBool(GetLine(bAdventure, iPos)) '?
                    Adventure.EnableMenu = Not CBool(GetLine(bAdventure, iPos)) '?
                    bSound = CBool(GetLine(bAdventure, iPos))
                    bGraphics = CBool(GetLine(bAdventure, iPos))
                    For i As Integer = 0 To 1
                        If bSound Then
                            sFilename = GetLine(bAdventure, iPos) ' Filename
                            If sFilename <> "" Then
                                Dim sLoop As String = ""
                                If sFilename.EndsWith("##") Then
                                    sLoop = " loop=Y"
                                    sFilename = sFilename.Substring(0, sFilename.Length - 2)
                                End If
                                If i = 0 Then Adventure.Introduction(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & Adventure.Introduction(0).Description
                                If i = 1 Then Adventure.WinningText(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & Adventure.WinningText(0).Description
                            End If
                            If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                            If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                        End If
                        If bGraphics Then
                            sFilename = GetLine(bAdventure, iPos) ' Filename
                            If sFilename <> "" Then
                                If i = 0 Then Adventure.Introduction(0).Description = "<img src=""" & sFilename & """>" & Adventure.Introduction(0).Description
                                If i = 1 Then Adventure.WinningText(0).Description = "<img src=""" & sFilename & """>" & Adventure.WinningText(0).Description
                            End If
                            If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                            If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, True))
                        End If
                    Next
                    If v >= 4 Then
                        GetLine(bAdventure, iPos) ' Enable Panel
                        Adventure.sUserStatus = GetLine(bAdventure, iPos) ' Panel Text
                    End If
                    GetLine(bAdventure, iPos) ' Size ratio
                    GetLine(bAdventure, iPos) ' Weight ratio
                Else
                    .ShowFirstRoom = False
                End If

                If v >= 4 Then GetLine(bAdventure, iPos) ' ?


                '----------------------------------------------------------------------------------
                ' Locations                
                '----------------------------------------------------------------------------------

                Dim iNumLocations As Integer = CInt(GetLine(bAdventure, iPos))
                Dim iX As Integer = 3
                If v < 4 Then iX = 2
                Dim iLocations(iNumLocations, 11, iX) As Integer ' Temp Store
                Dim iLoc As Integer
                Dim colNewLocs As New Collection
                For iLoc = 1 To iNumLocations
                    Dim Location As New clsLocation
                    With Location
                        '.Key = "Location" & iLoc.ToString
                        Dim sKey As String = "Location" & iLoc.ToString
                        If a.htblLocations.ContainsKey(sKey) Then
                            While a.htblLocations.ContainsKey(sKey)
                                sKey = IncrementKey(sKey)
                            End While
                        End If
                        .Key = sKey
                        colNewLocs.Add(sKey)
                        .ShortDescription = New Description(GetLine(bAdventure, iPos))
                        .LongDescription = New Description(ConvText(GetLine(bAdventure, iPos)))
                        If v < 4 Then
                            Dim srdesc1 As String = GetLine(bAdventure, iPos) '?
                            If srdesc1 <> "" Then
                                Dim sd As New SingleDescription
                                sd.Description = ConvText(srdesc1)
                                sd.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartAfterDefaultDescription
                                sd.Restrictions.BracketSequence = "#"
                                .LongDescription.Add(sd)
                            End If
                        End If
                        For i As Integer = 0 To iCompassPoints - 1
                            iLocations(iLoc, i, 0) = SafeInt(GetLine(bAdventure, iPos)) ' Rooms
                            If iLocations(iLoc, i, 0) <> 0 Then
                                iLocations(iLoc, i, 1) = SafeInt(GetLine(bAdventure, iPos)) ' Tasks
                                iLocations(iLoc, i, 2) = SafeInt(GetLine(bAdventure, iPos)) ' Completed
                                If v >= 4 Then iLocations(iLoc, i, 3) = SafeInt(GetLine(bAdventure, iPos)) ' Mode                                
                            End If
                        Next
                        If v < 4 Then
                            Dim srdesc2_0 As String = GetLine(bAdventure, iPos) '?
                            Dim irtaskno2_0 As Integer = SafeInt(GetLine(bAdventure, iPos))
                            If srdesc2_0 <> "" OrElse irtaskno2_0 > 0 Then
                                Dim sd As New SingleDescription
                                sd.Description = ConvText(srdesc2_0)
                                sd.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartAfterDefaultDescription
                                If irtaskno2_0 > 0 Then
                                    Dim rest As New clsRestriction
                                    rest.eType = clsRestriction.RestrictionTypeEnum.Task
                                    rest.sKey1 = "Task" & irtaskno2_0
                                    rest.eMust = clsRestriction.MustEnum.Must
                                    rest.eTask = clsRestriction.TaskEnum.Complete
                                    sd.Restrictions.Add(rest)
                                End If
                                sd.Restrictions.BracketSequence = "#"
                                .LongDescription.Add(sd)
                            End If
                            Dim srdesc2_1 As String = GetLine(bAdventure, iPos) '?
                            Dim irtaskno2_1 As Integer = SafeInt(GetLine(bAdventure, iPos))
                            If srdesc2_1 <> "" OrElse irtaskno2_1 > 0 Then
                                Dim sd As New SingleDescription
                                sd.Description = ConvText(srdesc2_1)
                                sd.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartAfterDefaultDescription
                                If irtaskno2_1 > 0 Then
                                    Dim rest As New clsRestriction
                                    rest.eType = clsRestriction.RestrictionTypeEnum.Task
                                    rest.sKey1 = "Task" & irtaskno2_1
                                    rest.eMust = clsRestriction.MustEnum.Must
                                    rest.eTask = clsRestriction.TaskEnum.Complete
                                    sd.Restrictions.Add(rest)
                                End If
                                sd.Restrictions.BracketSequence = "#"
                                .LongDescription.Add(sd)
                            End If
                            Dim irobject As Integer = SafeInt(GetLine(bAdventure, iPos))
                            Dim srdesc3 As String = GetLine(bAdventure, iPos) '?
                            Dim irhideob As Integer = SafeInt(GetLine(bAdventure, iPos))
                            If srdesc3 <> "" Then
                                Dim sd As New SingleDescription
                                sd.Description = ConvText(srdesc3)
                                sd.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartDescriptionWithThis
                                Dim rest As New clsRestriction
                                rest.eType = clsRestriction.RestrictionTypeEnum.Character
                                rest.sKey1 = "Player"
                                Select Case CInt(irhideob / 10)
                                    Case 0
                                        rest.eMust = clsRestriction.MustEnum.MustNot
                                        rest.eCharacter = clsRestriction.CharacterEnum.BeHoldingObject
                                    Case 1
                                        rest.eMust = clsRestriction.MustEnum.Must
                                        rest.eCharacter = clsRestriction.CharacterEnum.BeHoldingObject
                                    Case 2
                                        rest.eMust = clsRestriction.MustEnum.MustNot
                                        rest.eCharacter = clsRestriction.CharacterEnum.BeWearingObject
                                    Case 3
                                        rest.eMust = clsRestriction.MustEnum.Must
                                        rest.eCharacter = clsRestriction.CharacterEnum.BeWearingObject
                                    Case 4
                                        rest.eMust = clsRestriction.MustEnum.MustNot
                                        rest.eCharacter = clsRestriction.CharacterEnum.BeInSameLocationAsObject
                                    Case 5
                                        rest.eMust = clsRestriction.MustEnum.Must
                                        rest.eCharacter = clsRestriction.CharacterEnum.BeInSameLocationAsObject
                                End Select
                                rest.sKey2 = irobject.ToString ' Needs to be converted once we've loaded objects
                                sd.Restrictions.Add(rest)
                                sd.Restrictions.BracketSequence = "#"
                                .LongDescription.Add(sd)
                            End If
                        End If
                        If bSound Then
                            sFilename = GetLine(bAdventure, iPos) ' Filename
                            If sFilename <> "" Then
                                Dim sLoop As String = ""
                                If sFilename.EndsWith("##") Then
                                    sLoop = " loop=Y"
                                    sFilename = sFilename.Substring(0, sFilename.Length - 2)
                                End If
                                .LongDescription(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & .LongDescription(0).Description
                            End If
                            If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                            If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                        End If
                        If bGraphics Then
                            sFilename = GetLine(bAdventure, iPos) ' Filename
                            If sFilename <> "" Then
                                .LongDescription(0).Description = "<img src=""" & sFilename & """>" & .LongDescription(0).Description
                            End If
                            If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                            If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                        End If
                        Dim iNumAltDesc As Integer
                        If v >= 4 Then
                            iNumAltDesc = CInt(GetLine(bAdventure, iPos))
                        Else
                            iNumAltDesc = 4
                        End If
                        For iAlt As Integer = 0 To iNumAltDesc - 1
                            Dim rest As clsRestriction = Nothing
                            Dim sd As SingleDescription = Nothing
                            Dim iTaskObPlayer As Integer
                            If v >= 4 Then
                                sd = New SingleDescription

                                sd.Description = ConvText(GetLine(bAdventure, iPos)) ' Description
                                rest = New clsRestriction
                                iTaskObPlayer = CInt(GetLine(bAdventure, iPos)) ' Options
                                Select Case iTaskObPlayer
                                    Case 0 ' Task
                                        rest.eType = clsRestriction.RestrictionTypeEnum.Task
                                    Case 1 ' Object
                                        rest.eType = clsRestriction.RestrictionTypeEnum.Object
                                    Case 2 ' Player
                                        rest.eType = clsRestriction.RestrictionTypeEnum.Character
                                        rest.sKey1 = "%Player%"
                                End Select
                            End If
                            If bSound Then
                                sFilename = GetLine(bAdventure, iPos) ' Filename
                                If sFilename <> "" Then
                                    Dim sLoop As String = ""
                                    If sFilename.EndsWith("##") Then
                                        sLoop = " loop=Y"
                                        sFilename = sFilename.Substring(0, sFilename.Length - 2)
                                    End If
                                    If v < 4 Then
                                        .LongDescription(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & .LongDescription(0).Description
                                    Else
                                        sd.Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & sd.Description
                                    End If
                                End If
                                If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                            End If
                            If bGraphics Then
                                sFilename = GetLine(bAdventure, iPos) ' Filename
                                If sFilename <> "" Then
                                    If v < 4 Then
                                        .LongDescription(0).Description = "<img src=""" & sFilename & """>" & .LongDescription(0).Description
                                    Else
                                        sd.Description = "<img src=""" & sFilename & """>" & sd.Description
                                    End If
                                End If
                                If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, True))
                            End If

                            If v >= 4 Then
                                rest.oMessage = New Description(ConvText(GetLine(bAdventure, iPos))) ' Description
                                iTaskObPlayer = CInt(GetLine(bAdventure, iPos)) ' Options
                                Select Case rest.eType
                                    Case clsRestriction.RestrictionTypeEnum.Task
                                        rest.sKey1 = "Task" & iTaskObPlayer
                                    Case clsRestriction.RestrictionTypeEnum.Object
                                        rest.sKey1 = "Object" & iTaskObPlayer
                                        rest.eMust = clsRestriction.MustEnum.Must
                                        rest.eObject = clsRestriction.ObjectEnum.BeInState
                                    Case clsRestriction.RestrictionTypeEnum.Character
                                        Select Case iTaskObPlayer
                                            Case 0 ' is not holding
                                                rest.eCharacter = clsRestriction.CharacterEnum.BeHoldingObject
                                                rest.eMust = clsRestriction.MustEnum.MustNot
                                            Case 1 ' is holding
                                                rest.eCharacter = clsRestriction.CharacterEnum.BeHoldingObject
                                                rest.eMust = clsRestriction.MustEnum.Must
                                            Case 2 ' is not wearing
                                                rest.eCharacter = clsRestriction.CharacterEnum.BeWearingObject
                                                rest.eMust = clsRestriction.MustEnum.MustNot
                                            Case 3 ' is wearing
                                                rest.eCharacter = clsRestriction.CharacterEnum.BeHoldingObject
                                                rest.eMust = clsRestriction.MustEnum.Must
                                            Case 4 ' is not same room as
                                                rest.eCharacter = clsRestriction.CharacterEnum.BeInSameLocationAsObject
                                                rest.eMust = clsRestriction.MustEnum.MustNot
                                            Case 5 ' is in same room as
                                                rest.eCharacter = clsRestriction.CharacterEnum.BeInSameLocationAsObject
                                                rest.eMust = clsRestriction.MustEnum.Must
                                        End Select
                                End Select
                                If bSound Then
                                    sFilename = GetLine(bAdventure, iPos) ' Filename
                                    Dim sLoop As String = ""
                                    If sFilename.EndsWith("##") Then
                                        sLoop = " loop=Y"
                                        sFilename = sFilename.Substring(0, sFilename.Length - 2)
                                    End If
                                    If sFilename <> "" Then sd.Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & sd.Description
                                    iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                    If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                                End If
                                If bGraphics Then
                                    sFilename = GetLine(bAdventure, iPos) ' Filename
                                    If sFilename <> "" Then sd.Description = "<img src=""" & sFilename & """>" & sd.Description
                                    iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                    If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, True))
                                End If
                                GetLine(bAdventure, iPos) ' Hideobs
                                Dim sNewShort As String = GetLine(bAdventure, iPos) ' New Short Desc
                                iTaskObPlayer = CInt(GetLine(bAdventure, iPos)) ' Options
                                Select Case rest.eType
                                    Case clsRestriction.RestrictionTypeEnum.Task
                                        If iTaskObPlayer = 0 Then rest.eMust = clsRestriction.MustEnum.Must Else rest.eMust = clsRestriction.MustEnum.MustNot
                                        rest.eTask = clsRestriction.TaskEnum.Complete
                                    Case clsRestriction.RestrictionTypeEnum.Object
                                        rest.sKey2 = iTaskObPlayer.ToString
                                    Case clsRestriction.RestrictionTypeEnum.Character
                                        rest.sKey2 = iTaskObPlayer.ToString
                                End Select
                                Dim iDisplayWhen As Integer = CInt(GetLine(bAdventure, iPos))
                                sd.eDisplayWhen = CType(iDisplayWhen, SingleDescription.DisplayWhenEnum) ' Show When                            
                                If Not (rest.eType = clsRestriction.RestrictionTypeEnum.Task AndAlso rest.sKey1 = "Task0") Then sd.Restrictions.Add(rest)
                                sd.Restrictions.BracketSequence = "#"
                                .LongDescription.Add(sd)
                                If sNewShort <> "" Then
                                    Dim sdShort As SingleDescription = sd.CloneMe
                                    sdShort.Description = sNewShort
                                    sdShort.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartDescriptionWithThis
                                    .ShortDescription.Add(sdShort)
                                End If
                            End If
                        Next
                        If Adventure.Enabled(clsAdventure.EnabledOptionEnum.Map) Then .HideOnMap = CBool(SafeInt(GetLine(bAdventure, iPos))) ' No Map
                    End With
                    .htblLocations.Add(Location, Location.Key)
NextLoc:
                Next


                '----------------------------------------------------------------------------------
                ' Objects
                '----------------------------------------------------------------------------------

                Dim dictDodgyStates As New Dictionary(Of String, String)
                Dim iNumObjects As Integer = CInt(GetLine(bAdventure, iPos))
                Dim colNewObs As New Collection
                For iObj As Integer = 1 To iNumObjects
                    Dim NewObject As New clsObject
                    With NewObject
                        '.Key = "Object" & iObj.ToString
                        Dim sKey As String = "Object" & iObj.ToString
                        If a.htblObjects.ContainsKey(sKey) Then
                            While a.htblObjects.ContainsKey(sKey)
                                sKey = IncrementKey(sKey)
                            End While
                        End If
                        .Key = sKey
                        colNewObs.Add(sKey)
                        .Prefix = GetLine(bAdventure, iPos)
                        ConvertPrefix(.Article, .Prefix)
                        .arlNames.Add(GetLine(bAdventure, iPos))
                        If v < 4 Then
                            Dim sAlias As String = GetLine(bAdventure, iPos)
                            If sAlias <> "" Then .arlNames.Add(sAlias)
                        Else
                            Dim iNumAliases As Integer = CInt(GetLine(bAdventure, iPos))
                            For iAlias As Integer = 1 To iNumAliases
                                .arlNames.Add(GetLine(bAdventure, iPos))
                            Next
                        End If

                        Dim sod As New clsProperty
                        sod = Adventure.htblAllProperties("StaticOrDynamic").Copy
                        .AddProperty(sod)
                        .IsStatic = SafeBool(GetLine(bAdventure, iPos))

                        .Description = New Description(ConvText(GetLine(bAdventure, iPos)))
                        Dim cObjectLocation As New clsObjectLocation
                        If .IsStatic Then
                            GetLine(bAdventure, iPos) ' Not needed here?
                            Dim sl As New clsProperty
                            sl = Adventure.htblAllProperties("StaticLocation").Copy
                            .AddProperty(sl)
                        Else
                            Dim dl As New clsProperty
                            dl = Adventure.htblAllProperties("DynamicLocation").Copy
                            .AddProperty(dl)

                            Dim iv4Loc As Integer = CInt(GetLine(bAdventure, iPos))
                            If v < 3.9 AndAlso iv4Loc > 2 Then iv4Loc -= 1

                            Select Case iv4Loc
                                Case 0
                                    cObjectLocation.DynamicExistWhere = clsObjectLocation.DynamicExistsWhereEnum.Hidden
                                Case 1
                                    cObjectLocation.DynamicExistWhere = clsObjectLocation.DynamicExistsWhereEnum.HeldByCharacter
                                Case 2
                                    cObjectLocation.DynamicExistWhere = clsObjectLocation.DynamicExistsWhereEnum.InObject
                                Case 3
                                    cObjectLocation.DynamicExistWhere = clsObjectLocation.DynamicExistsWhereEnum.OnObject
                                Case iNumLocations + 4
                                    cObjectLocation.DynamicExistWhere = clsObjectLocation.DynamicExistsWhereEnum.WornByCharacter
                                Case Else
                                    cObjectLocation.DynamicExistWhere = clsObjectLocation.DynamicExistsWhereEnum.InLocation
                                    cObjectLocation.Key = "Location" & iv4Loc - 3 + iStartLocations
                                    Dim p As New clsProperty
                                    p = Adventure.htblAllProperties("InLocation").Copy
                                    .AddProperty(p)
                            End Select
                        End If

                        Dim sTaskKey As String = "Task" & GetLine(bAdventure, iPos)
                        Dim bTaskState As Boolean = CBool(GetLine(bAdventure, iPos))
                        Dim sDescription As String = ConvText(GetLine(bAdventure, iPos))
                        If sTaskKey <> "Task0" Then
                            Dim sd As New SingleDescription
                            sd.Description = sDescription
                            Dim rest As New clsRestriction
                            rest.eType = clsRestriction.RestrictionTypeEnum.Task
                            rest.eMust = CType(IIf(bTaskState, clsRestriction.MustEnum.MustNot, clsRestriction.MustEnum.Must), clsRestriction.MustEnum)
                            rest.eTask = clsRestriction.TaskEnum.Complete
                            rest.sKey1 = sTaskKey
                            sd.Restrictions.Add(rest)
                            sd.Restrictions.BracketSequence = "#"
                            sd.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartDescriptionWithThis
                            .Description.Add(sd)
                        End If

                        Dim StaticLoc As New clsObjectLocation
                        If .IsStatic Then
                            StaticLoc.StaticExistWhere = CType(GetLine(bAdventure, iPos), clsObjectLocation.StaticExistsWhereEnum)

                            Select Case StaticLoc.StaticExistWhere
                                Case clsObjectLocation.StaticExistsWhereEnum.NoRooms

                                Case clsObjectLocation.StaticExistsWhereEnum.SingleLocation
                                    StaticLoc.Key = "Location" & GetLine(bAdventure, iPos) ' StaticKey
                                    Dim pLoc As New clsProperty
                                    pLoc = Adventure.htblAllProperties("AtLocation").Copy
                                    .AddProperty(pLoc)
                                Case clsObjectLocation.StaticExistsWhereEnum.LocationGroup
                                    Dim salRooms As New StringArrayList
                                    For iRoom As Integer = 0 To iNumLocations
                                        If CBool(GetLine(bAdventure, iPos)) Then salRooms.Add("Location" & iRoom) ' StaticLoc.StaticKey = "Location" & iRoom ' TODO - Generate a roomgroup and assign that
                                    Next
                                    StaticLoc.Key = GetRoomGroupFromList(iPos, salRooms, "object '" & .FullName & "'").Key ' StaticKey
                                    Dim pLG As New clsProperty
                                    pLG = Adventure.htblAllProperties("AtLocationGroup").Copy
                                    .AddProperty(pLG)
                                Case clsObjectLocation.StaticExistsWhereEnum.AllRooms

                                Case clsObjectLocation.StaticExistsWhereEnum.PartOfCharacter
                                    ' Key defined later
                                Case clsObjectLocation.StaticExistsWhereEnum.PartOfObject
                                    ' Key defined later
                            End Select
                        End If


                        If CBool(GetLine(bAdventure, iPos)) Then ' container
                            Dim c As New clsProperty
                            c = Adventure.htblAllProperties("Container").Copy
                            c.Selected = True
                            .AddProperty(c)
                        End If
                        If CBool(GetLine(bAdventure, iPos)) Then ' surface
                            Dim s As New clsProperty
                            s = Adventure.htblAllProperties("Surface").Copy
                            s.Selected = True
                            .AddProperty(s)
                        End If
                        Dim iCapacity As Integer = CInt(GetLine(bAdventure, iPos)) ' Num Holds
                        If v < 3.9 Then iCapacity = iCapacity * 100 + 2
                        If .IsContainer Then
                            Dim iCapacity2 As Integer = iCapacity Mod 10
                            iCapacity -= iCapacity2
                            iCapacity = CInt((iCapacity / 10) * (3 ^ iCapacity2))
                            If Adventure.htblObjectProperties.ContainsKey("Capacity") Then
                                Dim p As New clsProperty
                                p = Adventure.htblObjectProperties("Capacity").Copy
                                p.Selected = True
                                p.Value = iCapacity.ToString
                                .AddProperty(p)
                            End If
                        End If

                        If Not .IsStatic Then
                            If CBool(GetLine(bAdventure, iPos)) Then  ' wearable
                                Dim w As New clsProperty
                                w = Adventure.htblAllProperties("Wearable").Copy
                                w.Selected = True
                                .AddProperty(w)
                            End If
                            Dim iSize As Integer = CInt(GetLine(bAdventure, iPos)) ' weight                            
                            Dim iWeight As Integer = iSize Mod 10
                            If v < 3.9 Then
                                Select Case iSize
                                    Case 0
                                        iWeight = 22
                                    Case 1
                                        iWeight = 23
                                    Case 2
                                        iWeight = 24
                                    Case 3
                                        iWeight = 32
                                    Case 4
                                        iWeight = 42
                                End Select
                            Else
                                iSize -= iWeight
                            End If

                            Dim p As New clsProperty
                            If Adventure.htblObjectProperties.ContainsKey("Size") Then
                                p = Adventure.htblObjectProperties("Size").Copy
                                p.Selected = True
                                p.Value = (3 ^ (iSize / 10)).ToString
                                .AddProperty(p)
                            End If

                            If Adventure.htblObjectProperties.ContainsKey("Weight") Then
                                p = New clsProperty
                                p = Adventure.htblObjectProperties("Weight").Copy
                                p.Selected = True
                                p.Value = (3 ^ iWeight).ToString
                                .AddProperty(p)
                            End If

                            Dim iParent As Integer = CInt(GetLine(bAdventure, iPos))
                            Select Case cObjectLocation.DynamicExistWhere
                                Case clsObjectLocation.DynamicExistsWhereEnum.HeldByCharacter
                                    p = New clsProperty
                                    p = Adventure.htblAllProperties("HeldByWho").Copy
                                    .AddProperty(p)

                                    If iParent = 0 Then
                                        cObjectLocation.Key = "%Player%"
                                    Else
                                        cObjectLocation.Key = "Character" & iParent
                                    End If
                                    .Location = cObjectLocation
                                Case clsObjectLocation.DynamicExistsWhereEnum.WornByCharacter
                                    If iParent = 0 Then
                                        cObjectLocation.Key = "%Player%"
                                    Else
                                        cObjectLocation.Key = "Character" & iParent
                                    End If
                                    p = New clsProperty
                                    p = Adventure.htblAllProperties("WornByWho").Copy
                                    .AddProperty(p)
                                    p.Value = cObjectLocation.Key
                                    .Location = cObjectLocation
                                Case clsObjectLocation.DynamicExistsWhereEnum.InObject
                                    .Location = cObjectLocation
                                    p = New clsProperty
                                    p = Adventure.htblAllProperties("InsideWhat").Copy
                                    .AddProperty(p)
                                    p.Value = iParent.ToString
                                Case clsObjectLocation.DynamicExistsWhereEnum.OnObject
                                    .Location = cObjectLocation
                                    p = New clsProperty
                                    p = Adventure.htblAllProperties("OnWhat").Copy
                                    .AddProperty(p)
                                    p.Value = iParent.ToString
                                Case Else
                                    .Location = cObjectLocation
                            End Select
                        End If
                        If .IsStatic AndAlso StaticLoc.StaticExistWhere = clsObjectLocation.StaticExistsWhereEnum.PartOfCharacter Then
                            Dim iChar As Integer = CInt(GetLine(bAdventure, iPos))
                            If iChar = 0 Then
                                StaticLoc.Key = "%Player%" ' StaticKey
                            Else
                                StaticLoc.Key = "Character" & iChar ' StaticKey
                            End If
                            Dim c As New clsProperty
                            c = Adventure.htblAllProperties("PartOfWho").Copy
                            c.StringData = New Description(ConvText(StaticLoc.Key)) ' StaticKey
                            .AddProperty(c)
                        End If
                        If .IsStatic Then .Move(StaticLoc)

                        Dim iOpenableLockable As Integer = CInt(GetLine(bAdventure, iPos))
                        If v < 4 AndAlso iOpenableLockable > 1 Then iOpenableLockable = 11 - iOpenableLockable
                        ' 0 = Not openable
                        ' 5 = Openable, open
                        ' 6 = Openable, closed
                        ' 7 = Openable, locked
                        If iOpenableLockable > 0 Then
                            Dim op As New clsProperty
                            op = Adventure.htblAllProperties("Openable").Copy
                            op.Selected = True
                            '.htblActualProperties.Add(op)
                            '.bCalculatedGroups = False
                            .AddProperty(op)

                            Dim pOS As New clsProperty
                            pOS = Adventure.htblAllProperties("OpenStatus").Copy
                            pOS.Selected = True
                            If iOpenableLockable = 5 Then
                                pOS.Value = "Open"
                            Else
                                pOS.Value = "Closed"
                            End If
                            .AddProperty(pOS)

                            If v >= 4 Then
                                Dim iKey As Integer = CInt(GetLine(bAdventure, iPos))
                                If iKey > -1 Then
                                    Dim pLk As New clsProperty
                                    pLk = Adventure.htblAllProperties("Lockable").Copy
                                    pLk.Selected = True
                                    .AddProperty(pLk)

                                    Dim pKey As New clsProperty
                                    pKey = Adventure.htblAllProperties("LockKey").Copy
                                    pKey.Selected = True
                                    pKey.Value = CStr(iKey) ' "Object" & iKey + iStartObs
                                    .AddProperty(pKey)

                                    Dim pLS As New clsProperty
                                    pLS = Adventure.htblAllProperties("LockStatus").Copy
                                    pLS.Selected = True
                                    If iOpenableLockable = 7 Then pOS.Value = "Locked"
                                    .AddProperty(pLS)
                                End If
                            End If
                        End If

                        Dim iSitStandLie As Integer = CInt(GetLine(bAdventure, iPos))  ' Sittable
                        If iSitStandLie = 1 OrElse iSitStandLie = 3 Then .IsSittable = True
                        .IsStandable = .IsSittable
                        If iSitStandLie = 2 OrElse iSitStandLie = 3 Then .IsLieable = True
                        If Not .IsStatic Then GetLine(bAdventure, iPos) ' edible

                        If CBool(GetLine(bAdventure, iPos)) Then
                            Dim r As New clsProperty
                            r = Adventure.htblAllProperties("Readable").Copy
                            r.Selected = True
                            .AddProperty(r)
                        End If
                        If .Readable Then
                            Dim sReadText As String = GetLine(bAdventure, iPos)
                            If sReadText <> "" Then
                                Dim r As New clsProperty
                                r = Adventure.htblAllProperties("ReadText").Copy
                                r.Selected = True
                                .AddProperty(r)
                                .ReadText = sReadText
                            End If
                        End If

                        If Not .IsStatic Then GetLine(bAdventure, iPos) ' weapon

                        If v >= 4 Then
                            Dim iState As Integer = CInt(GetLine(bAdventure, iPos))
                            If iState > 0 Then
                                Dim sStates As String = GetLine(bAdventure, iPos)
                                Dim arlStates As New StringArrayList
                                For Each sState As String In sStates.Split("|"c)
                                    arlStates.Add(ToProper(sState))
                                Next
                                Dim sPKey As String = FindProperty(arlStates)
                                Dim s As New clsProperty
                                If sPKey Is Nothing Then
                                    s.Type = clsProperty.PropertyTypeEnum.StateList
                                    s.Description = "Object can be " & sStates.Replace("|", " or ")
                                    s.Key = sStates
                                    s.arlStates = arlStates
                                    s.Value = arlStates(iState - 1)
                                    s.PropertyOf = clsProperty.PropertyOfEnum.Objects
                                    Adventure.htblAllProperties.Add(s.Copy)
                                Else
                                    s = Adventure.htblAllProperties(sPKey).Copy
                                    s.Value = arlStates(iState - 1)
                                    If sStates <> sPKey Then
                                        ' Hmm, the states are not in the same order as before
                                        ' This can cause problems if restrictions/actions use the state index                                        
                                        dictDodgyStates.Add(.Key, sStates)
                                    End If
                                    sStates = sPKey
                                End If
                                s.Selected = True
                                .AddProperty(s)
                                salWithStates.Add(.Key)

                                Dim bShowState As Boolean = CBool(GetLine(bAdventure, iPos)) ' showstate
                                If bShowState Then
                                    .Description(0).Description &= "  " & .Key & ".Name is %LCase[" & .Key & "." & sStates & "]%."
                                End If
                            End If
                            Dim bSpecificallyList As Boolean = CBool(GetLine(bAdventure, iPos)) ' showhide
                            If .IsStatic Then
                                .ExplicitlyList = bSpecificallyList
                            Else
                                .ExplicitlyExclude = bSpecificallyList
                            End If
                        End If
                        ' GSFX
                        For i As Integer = 0 To 1
                            If bSound Then
                                sFilename = GetLine(bAdventure, iPos) ' Filename
                                If sFilename <> "" Then
                                    Dim sLoop As String = ""
                                    If sFilename.EndsWith("##") Then
                                        sLoop = " loop=Y"
                                        sFilename = sFilename.Substring(0, sFilename.Length - 2)
                                    End If
                                    .Description(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & .Description(0).Description
                                End If
                                If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                            End If
                            If bGraphics Then
                                sFilename = GetLine(bAdventure, iPos) ' Filename
                                If sFilename <> "" Then
                                    .Description(0).Description = "<img src=""" & sFilename & """>" & .Description(0).Description
                                End If
                                If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, True))
                            End If
                        Next
                        ' Battle
                        If bBattleSystem Then
                            GetLine(bAdventure, iPos) ' Armour
                            GetLine(bAdventure, iPos) ' Hit Points
                            GetLine(bAdventure, iPos) ' Hit Method
                            If v >= 4 Then GetLine(bAdventure, iPos) ' Accuracy
                        End If
                        If v >= 4 Then
                            Dim sSpecialList As String = GetLine(bAdventure, iPos) ' alsohere
                            If sSpecialList <> "" Then
                                Dim r As New clsProperty
                                If NewObject.IsStatic Then
                                    r = Adventure.htblAllProperties("ListDescription").Copy
                                Else
                                    r = Adventure.htblAllProperties("ListDescriptionDynamic").Copy
                                End If
                                r.Selected = True
                                .AddProperty(r)
                                .ListDescription = sSpecialList
                            End If
                            Dim s2 As String = GetLine(bAdventure, iPos) ' initial
                        End If
                    End With

                    .htblObjects.Add(NewObject, NewObject.Key)
                Next

                ' Sort out object keys
                If Adventure.Player.Location.ExistWhere = clsCharacterLocation.ExistsWhereEnum.OnObject Then
                    Select Case Adventure.Player.Location.Position
                        Case clsCharacterLocation.PositionEnum.Standing
                            Adventure.Player.Location.Key = GetObKey(CInt(Adventure.Player.Location.Key) - 1, ComboEnum.Standable)
                        Case clsCharacterLocation.PositionEnum.Sitting
                            Adventure.Player.Location.Key = GetObKey(CInt(Adventure.Player.Location.Key) - 1, ComboEnum.Sittable)
                        Case clsCharacterLocation.PositionEnum.Lying
                            Adventure.Player.Location.Key = GetObKey(CInt(Adventure.Player.Location.Key) - 1, ComboEnum.Lieable)
                    End Select
                End If
                For Each sLoc As String In colNewLocs
                    Dim loc As clsLocation = a.htblLocations(sLoc)
                    Dim listDescriptions As New List(Of SingleDescription)

                    For Each sd As SingleDescription In loc.ShortDescription
                        listDescriptions.Add(sd)
                    Next
                    For Each sd As SingleDescription In loc.LongDescription
                        listDescriptions.Add(sd)
                    Next

                    For Each sd As SingleDescription In listDescriptions
                        If sd.Restrictions.Count > 0 AndAlso IsNumeric(sd.Restrictions(0).sKey2) Then
                            If sd.Restrictions(0).eType = clsRestriction.RestrictionTypeEnum.Character Then
                                Select Case sd.Restrictions(0).eCharacter
                                    Case clsRestriction.CharacterEnum.BeInSameLocationAsObject, clsRestriction.CharacterEnum.BeHoldingObject
                                        sd.Restrictions(0).sKey2 = GetObKey(CInt(sd.Restrictions(0).sKey2) - 1, ComboEnum.Dynamic)
                                    Case clsRestriction.CharacterEnum.BeWearingObject
                                        sd.Restrictions(0).sKey2 = GetObKey(CInt(sd.Restrictions(0).sKey2) - 1, ComboEnum.Wearable)
                                End Select
                            ElseIf sd.Restrictions(0).eType = clsRestriction.RestrictionTypeEnum.Object Then
                                Dim ob As clsObject = Nothing
                                Dim sKey1 As String = sd.Restrictions(0).sKey1
                                Dim sKey2 As String = sd.Restrictions(0).sKey2
                                If sKey1 <> "ReferencedObject" Then ob = Adventure.htblObjects(sKey1)

                                Select Case CInt(sd.Restrictions(0).sKey2)
                                    Case 1
                                        If sKey1 = "ReferencedObject" OrElse ob.Openable Then
                                            sKey2 = "Open"
                                        Else
                                            For Each prop As clsProperty In ob.htblActualProperties.Values
                                                If prop.Key.IndexOf("|"c) > 0 Then
                                                    sKey2 = prop.arlStates(CInt(sd.Restrictions(0).sKey2) - 1)
                                                End If
                                            Next
                                        End If
                                    Case 2
                                        If sKey1 = "ReferencedObject" OrElse ob.Openable Then
                                            sKey2 = "Closed"
                                        Else
                                            For Each prop As clsProperty In ob.htblActualProperties.Values
                                                If prop.Key.IndexOf("|"c) > 0 Then
                                                    sKey2 = prop.arlStates(CInt(sd.Restrictions(0).sKey2) - 1)
                                                End If
                                            Next
                                        End If
                                    Case 3
                                        If sKey1 = "ReferencedObject" OrElse ob.Openable Then
                                            If sKey1 = "ReferencedObject" OrElse ob.Lockable Then
                                                sKey2 = "Locked"
                                            Else
                                                For Each prop As clsProperty In ob.htblActualProperties.Values
                                                    If prop.Key.IndexOf("|"c) > 0 Then
                                                        sKey2 = prop.arlStates(CInt(sd.Restrictions(0).sKey2) - 3)
                                                    End If
                                                Next
                                            End If
                                        Else
                                            For Each prop As clsProperty In ob.htblActualProperties.Values
                                                If prop.Key.IndexOf("|"c) > 0 Then
                                                    sKey2 = prop.arlStates(CInt(sd.Restrictions(0).sKey2) - 1)
                                                End If
                                            Next
                                        End If
                                    Case Else
                                        Dim iOffset As Integer = 0
                                        If sKey1 = "ReferencedObject" OrElse ob.Openable Then
                                            If sKey1 = "ReferencedObject" OrElse ob.Lockable Then iOffset = 4 Else iOffset = 3
                                        End If
                                        For Each prop As clsProperty In ob.htblActualProperties.Values
                                            If prop.Key.IndexOf("|"c) > 0 Then
                                                sKey2 = prop.arlStates(CInt(sd.Restrictions(0).sKey2) - iOffset)
                                            End If
                                        Next
                                End Select
                                sd.Restrictions(0).sKey2 = sKey2
                            End If
                        End If
                    Next
                Next
                For Each sOb As String In colNewObs
                    Dim ob As clsObject = a.htblObjects(sOb)
                    If ob.Lockable Then
                        ob.SetPropertyValue("LockKey", GetObKey(CInt(ob.GetPropertyValue("LockKey")) + iStartObs, ComboEnum.Dynamic))
                    End If
                    If ob.HasProperty("OnWhat") Then
                        ob.SetPropertyValue("OnWhat", GetObKey(CInt(ob.GetPropertyValue("OnWhat")), ComboEnum.Surface))
                    End If
                    If ob.HasProperty("InsideWhat") Then
                        ob.SetPropertyValue("InsideWhat", GetObKey(CInt(ob.GetPropertyValue("InsideWhat")), ComboEnum.Container))
                    End If
                Next

                ' Sort out location restrictions
                'For iLoc = 1 To .htblLocations.Count
                iLoc = 0
                For Each sLoc As String In colNewLocs
                    iLoc += 1
                    Dim loc As clsLocation = Adventure.htblLocations(sLoc)
                    For iDir As DirectionsEnum = DirectionsEnum.North To DirectionsEnum.NorthWest
                        If iLocations(iLoc, iDir, 0) > 0 Then
                            If iLocations(iLoc, iDir, 0) <= Adventure.htblLocations.Count Then
                                loc.arlDirections(iDir).LocationKey = "Location" & iLocations(iLoc, iDir, 0)
                            Else
                                loc.arlDirections(iDir).LocationKey = "Group" & iLocations(iLoc, iDir, 0) - Adventure.htblLocations.Count
                            End If
                            If iLocations(iLoc, iDir, 1) > 0 Then
                                Dim rest As New clsRestriction
                                If v < 4 OrElse iLocations(iLoc, iDir, 3) = 0 Then
                                    rest.eType = clsRestriction.RestrictionTypeEnum.Task
                                    rest.sKey1 = "Task" & iLocations(iLoc, iDir, 1) + iStartTask
                                    If iLocations(iLoc, iDir, 2) = 0 Then
                                        rest.eMust = clsRestriction.MustEnum.Must
                                    Else
                                        rest.eMust = clsRestriction.MustEnum.MustNot
                                    End If
                                    rest.eTask = clsRestriction.TaskEnum.Complete
                                Else
                                    rest.eType = clsRestriction.RestrictionTypeEnum.Property
                                    ' Filter on objects with state
                                    rest.sKey1 = "OpenStatus"
                                    rest.sKey2 = GetObKey(iLocations(iLoc, iDir, 1) - 1, ComboEnum.WithStateOrOpenable)
                                    rest.eMust = clsRestriction.MustEnum.Must
                                    Dim ob As clsObject = Adventure.htblObjects(rest.sKey2)
                                    If ob.Openable Then
                                        If iLocations(iLoc, iDir, 2) = 0 Then rest.StringValue = "Open"
                                        If iLocations(iLoc, iDir, 2) = 1 Then rest.StringValue = "Closed"
                                        If ob.Lockable AndAlso iLocations(iLoc, iDir, 2) = 2 Then rest.StringValue = "Locked"
                                    End If
                                End If
                                loc.arlDirections(iDir).Restrictions.Add(rest)
                                loc.arlDirections(iDir).Restrictions.BracketSequence = "#"
                            End If
                        End If
                    Next
                Next



                '----------------------------------------------------------------------------------
                ' Tasks                
                '----------------------------------------------------------------------------------

                Dim iNumTasks As Integer = CInt(GetLine(bAdventure, iPos))
                For iTask As Integer = 1 To iNumTasks
                    Dim NewTask As New clsTask
                    With NewTask
                        Dim sKey As String = "Task" & iTask.ToString
                        If a.htblTasks.ContainsKey(sKey) Then
                            While a.htblTasks.ContainsKey(sKey)
                                sKey = IncrementKey(sKey)
                            End While
                        End If
                        .Key = sKey
                        .Priority = iStartMaxPriority + iTask
                        Dim iNumCommands As Integer = CInt(GetLine(bAdventure, iPos))
                        If v < 4 Then iNumCommands += 1
                        For i As Integer = 1 To iNumCommands
                            Dim sCommand As String = GetLine(bAdventure, iPos)
                            ' Simplify Runner so it only has to deal with multiple, or specific refs
                            .arlCommands.Add(sCommand.Replace("%object%", "%object1%").Replace("%character%", "%character1%"))
                        Next
                        .TaskType = clsTask.TaskTypeEnum.System
                        For Each sCommand As String In .arlCommands
                            If Left(sCommand, 1) <> "#" Then
                                .TaskType = clsTask.TaskTypeEnum.General
                                Exit For
                            End If
                        Next
                        .Description = .arlCommands(0)
                        Dim sMessage0 As String = GetLine(bAdventure, iPos)
                        .CompletionMessage = New Description(ConvText(sMessage0))
                        Dim sMessage1 As String = GetLine(bAdventure, iPos)
                        Dim sMessage2 As String = GetLine(bAdventure, iPos)
                        Dim sMessage3 As String = GetLine(bAdventure, iPos)
                        Dim iShowRoom As Integer = SafeInt(GetLine(bAdventure, iPos))
                        If iShowRoom > 0 Then
                            If .CompletionMessage(0).Description <> "" Then .CompletionMessage(0).Description &= "  "
                            .CompletionMessage(0).Description &= "%DisplayLocation[Location" & iShowRoom & "]%"
                        End If
                        If sMessage3 <> "" AndAlso .CompletionMessage.ToString = "" Then .eDisplayCompletion = clsTask.BeforeAfterEnum.After Else .eDisplayCompletion = clsTask.BeforeAfterEnum.Before
                        If .eDisplayCompletion = clsTask.BeforeAfterEnum.Before And sMessage0.Contains("%") Then .eDisplayCompletion = clsTask.BeforeAfterEnum.After ' v4 didn't handle this properly, so any text was substituted afterwards anyway
                        If sMessage3 <> "" Then .CompletionMessage(0).Description = pSpace(.CompletionMessage.ToString) & sMessage3
                        If .CompletionMessage.ToString = "" Then .SpecificOverrideType = clsTask.SpecificOverrideTypeEnum.BeforeTextAndActions
                        ' Needs to be ContinueOnFail so that a failing task with output will be overridden by a lower priority succeeding task, as per v4
                        .ContinueToExecuteLowerPriority = False
                        .Repeatable = CBool(GetLine(bAdventure, iPos))
                        If v < 3.9 Then
                            GetLine(bAdventure, iPos) ' score
                            GetLine(bAdventure, iPos) ' upto
                            For i As Integer = 0 To 5
                                Dim move_nn_0 As Integer = CInt(GetLine(bAdventure, iPos)) ' move(nn, 0)
                                Dim move_nn_1 As Integer = CInt(GetLine(bAdventure, iPos)) ' move(nn, 1)
                                If v >= 3.8 Then GetLine(bAdventure, iPos) ' movemode(nn)
                            Next
                        End If
                        GetLine(bAdventure, iPos) ' Reversible - give warning if this is set
                        iNumCommands = CInt(GetLine(bAdventure, iPos))
                        If v < 4 Then iNumCommands += 1
                        For i As Integer = 1 To iNumCommands
                            '.arlReverseCommands.Add(
                            GetLine(bAdventure, iPos)
                        Next

                        If v < 3.9 Then
                            For i As Integer = 0 To 1
                                GetLine(bAdventure, iPos) ' wear(nn)
                            Next
                            For i As Integer = 0 To 3
                                GetLine(bAdventure, iPos) ' hold(nn)
                            Next
                            GetLine(bAdventure, iPos) ' TaskNo
                            GetLine(bAdventure, iPos) ' TaskDone
                            For i As Integer = 0 To 3
                                GetLine(bAdventure, iPos) ' elses(i)
                            Next
                            GetLine(bAdventure, iPos) ' hereornot
                            GetLine(bAdventure, iPos) ' who
                            GetLine(bAdventure, iPos) ' elses(4)
                            GetLine(bAdventure, iPos) ' obroom
                        End If

                        ' Convert rooms executable in into restrictions
                        ' If up to 3 rooms, add as seperate restrictions
                        ' If up to 3 away from all then add as separate restrictions
                        ' Otherwise, create a room group and have that as single restriction
                        '
                        Dim iDoWhere As Integer = SafeInt(GetLine(bAdventure, iPos)) ' 0=None, 1=Single, 2=Multiple, 3=All
                        If iDoWhere = 1 Then
                            .arlRestrictions.Add(LocRestriction("Location" & CInt(GetLine(bAdventure, iPos)) + 1 + iStartLocations, True))
                            .arlRestrictions.BracketSequence = "#"
                        End If
                        If iDoWhere = 2 Then
                            Dim bHere As Boolean
                            Dim iCount As Integer = 0
                            Dim salHere As New StringArrayList
                            Dim salNotHere As New StringArrayList
                            For i As Integer = 1 To iNumLocations
                                bHere = CBool(GetLine(bAdventure, iPos))
                                If bHere Then
                                    iCount += 1
                                    salHere.Add("Location" & i + iStartLocations)
                                Else
                                    salNotHere.Add("Location" & i + iStartLocations)
                                End If
                            Next
                            Select Case iCount
                                Case 2, 3
                                    For Each sLocKey As String In salHere
                                        .arlRestrictions.Add(LocRestriction(sLocKey, True))
                                    Next
                                    If iCount = 2 Then
                                        .arlRestrictions.BracketSequence = "(#O#)"
                                    Else
                                        .arlRestrictions.BracketSequence = "(#O#O#)"
                                    End If
                                Case iNumLocations - 1, iNumLocations - 2
                                    For Each sLocKey As String In salNotHere
                                        .arlRestrictions.Add(LocRestriction(sLocKey, False))
                                    Next
                                    If iCount = iNumLocations - 1 Then
                                        .arlRestrictions.BracketSequence = "#"
                                    Else
                                        .arlRestrictions.BracketSequence = "(#O#)"
                                    End If
                                Case Else
                                    .arlRestrictions.Add(LocRestriction(GetRoomGroupFromList(iPos, salHere, "task '" & .Description & "'").Key, True))
                                    .arlRestrictions.BracketSequence = "#"
                            End Select
                        End If

                        If v < 3.9 Then
                            GetLine(bAdventure, iPos) ' kills
                            GetLine(bAdventure, iPos) ' holding
                        End If

                        Dim sQuestion As String = GetLine(bAdventure, iPos)
                        If sQuestion <> "" Then
                            Dim NewHint As New clsHint
                            With NewHint
                                .Key = "Hint" & (Adventure.htblHints.Count + 1).ToString
                                .Question = sQuestion
                                .SubtleHint = New Description(ConvText(GetLine(bAdventure, iPos)))
                                .SledgeHammerHint = New Description(ConvText(GetLine(bAdventure, iPos)))
                            End With
                            Adventure.htblHints.Add(NewHint, NewHint.Key)
                        End If

                        If v < 3.9 Then
                            Dim iObStuff As Integer = CInt(GetLine(bAdventure, iPos))
                            If iObStuff > 0 Then
                                GetLine(bAdventure, iPos) ' obstuff(1)
                                GetLine(bAdventure, iPos) ' obstuff(2)
                                GetLine(bAdventure, iPos) ' elses(5)
                            End If
                            GetLine(bAdventure, iPos) ' winning

                        Else
                            Dim iNumRestriction As Integer = CInt(GetLine(bAdventure, iPos))
                            For i As Integer = 1 To iNumRestriction
                                Dim NewRestriction As New clsRestriction
                                With NewRestriction
                                    Dim iMode As Integer = CInt(GetLine(bAdventure, iPos))
                                    Dim iCombo0 As Integer = CInt(GetLine(bAdventure, iPos))
                                    Dim iCombo1 As Integer = CInt(GetLine(bAdventure, iPos))
                                    Dim iCombo2 As Integer
                                    If iMode = 0 OrElse iMode > 2 Then iCombo2 = CInt(GetLine(bAdventure, iPos))
                                    If v < 4 AndAlso iMode = 4 AndAlso iCombo0 > 0 Then iCombo0 += 1
                                    Dim sText As String = Nothing
                                    If v >= 4 Then
                                        If iMode = 4 Then sText = GetLine(bAdventure, iPos)
                                    End If
                                    Select Case iMode
                                        Case 0 ' Object Locations
                                            .eType = clsRestriction.RestrictionTypeEnum.Object
                                            Select Case iCombo0
                                                Case 0
                                                    .sKey1 = NOOBJECT
                                                Case 1
                                                    .sKey1 = ANYOBJECT
                                                Case 2
                                                    .sKey1 = "ReferencedObject"
                                                Case Else
                                                    .sKey1 = GetObKey(iCombo0 - 3, ComboEnum.Dynamic)
                                            End Select
                                            Select Case iCombo1
                                                Case 0, 6
                                                    .eObject = clsRestriction.ObjectEnum.BeAtLocation
                                                    If iCombo1 = 6 Then .eMust = clsRestriction.MustEnum.MustNot
                                                    If iCombo2 = 0 Then
                                                        .eObject = clsRestriction.ObjectEnum.BeHidden
                                                    Else
                                                        .sKey2 = "Location" & iCombo2
                                                    End If
                                                Case 1, 7
                                                    .eObject = clsRestriction.ObjectEnum.BeHeldByCharacter
                                                    If iCombo1 = 7 Then .eMust = clsRestriction.MustEnum.MustNot
                                                    Select Case iCombo2
                                                        Case 0
                                                            .sKey2 = "%Player%"
                                                        Case 1
                                                            .sKey2 = "ReferencedCharacter"
                                                        Case Else
                                                            .sKey2 = "Character" & iCombo2 - 1 + iStartChar
                                                    End Select
                                                Case 2, 8
                                                    .eObject = clsRestriction.ObjectEnum.BeWornByCharacter
                                                    If iCombo1 = 8 Then .eMust = clsRestriction.MustEnum.MustNot
                                                    Select Case iCombo2
                                                        Case 0
                                                            .sKey2 = "%Player%"
                                                        Case 1
                                                            .sKey2 = "ReferencedCharacter"
                                                        Case Else
                                                            .sKey2 = "Character" & iCombo2 - 1 + iStartChar
                                                    End Select
                                                Case 3, 9
                                                    .eObject = clsRestriction.ObjectEnum.BeVisibleToCharacter
                                                    If iCombo1 = 9 Then .eMust = clsRestriction.MustEnum.MustNot
                                                    Select Case iCombo2
                                                        Case 0
                                                            .sKey2 = "%Player%"
                                                        Case 1
                                                            .sKey2 = "ReferencedCharacter"
                                                        Case Else
                                                            .sKey2 = "Character" & iCombo2 - 1 + iStartChar
                                                    End Select
                                                Case 4, 10
                                                    .eObject = clsRestriction.ObjectEnum.BeInsideObject
                                                    If iCombo1 = 10 Then .eMust = clsRestriction.MustEnum.MustNot
                                                    Select Case iCombo2
                                                        Case 0
                                                            ' Nothing
                                                        Case Else
                                                            .sKey2 = GetObKey(iCombo2 - 1, ComboEnum.Container)
                                                    End Select
                                                Case 5, 11
                                                    .eObject = clsRestriction.ObjectEnum.BeOnObject
                                                    If iCombo1 = 11 Then .eMust = clsRestriction.MustEnum.MustNot
                                                    Select Case iCombo2
                                                        Case 0
                                                            ' Nothing
                                                        Case Else
                                                            .sKey2 = GetObKey(iCombo2 - 1, ComboEnum.Surface)
                                                    End Select
                                            End Select
                                        Case 1 ' Object status
                                            .eType = clsRestriction.RestrictionTypeEnum.Object
                                            If iCombo0 = 0 Then
                                                .sKey1 = "ReferencedObject"
                                            Else
                                                .sKey1 = GetObKey(iCombo0 - 1, ComboEnum.WithStateOrOpenable)
                                            End If
                                            .eMust = clsRestriction.MustEnum.Must
                                            .eObject = clsRestriction.ObjectEnum.BeInState
                                            Dim ob As clsObject = Nothing
                                            If .sKey1 <> "ReferencedObject" Then ob = Adventure.htblObjects(.sKey1)
                                            Select Case iCombo1
                                                Case 0
                                                    If .sKey1 = "ReferencedObject" OrElse ob.Openable Then
                                                        .sKey2 = "Open"
                                                    Else
                                                        For Each prop As clsProperty In ob.htblActualProperties.Values
                                                            If prop.Key.IndexOf("|"c) > 0 Then
                                                                If dictDodgyStates.ContainsKey(.sKey1) Then
                                                                    Dim sIntended As String = dictDodgyStates(.sKey1).Split("|"c)(iCombo1)
                                                                    For Each state As String In prop.arlStates
                                                                        If sIntended = state.ToLower Then
                                                                            .sKey2 = state
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                Else
                                                                    .sKey2 = prop.arlStates(iCombo1)
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                Case 1
                                                    If .sKey1 = "ReferencedObject" OrElse ob.Openable Then
                                                        .sKey2 = "Closed"
                                                    Else
                                                        For Each prop As clsProperty In ob.htblActualProperties.Values
                                                            If prop.Key.IndexOf("|"c) > 0 Then
                                                                If dictDodgyStates.ContainsKey(.sKey1) Then
                                                                    Dim sIntended As String = dictDodgyStates(.sKey1).Split("|"c)(iCombo1)
                                                                    For Each state As String In prop.arlStates
                                                                        If sIntended = state.ToLower Then
                                                                            .sKey2 = state
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                Else
                                                                    .sKey2 = prop.arlStates(iCombo1)
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                Case 2
                                                    If .sKey1 = "ReferencedObject" OrElse ob.Openable Then
                                                        If .sKey1 = "ReferencedObject" OrElse ob.Lockable Then
                                                            .sKey2 = "Locked"
                                                        Else
                                                            For Each prop As clsProperty In ob.htblActualProperties.Values
                                                                If prop.Key.IndexOf("|"c) > 0 Then
                                                                    If dictDodgyStates.ContainsKey(.sKey1) Then
                                                                        Dim sIntended As String = dictDodgyStates(.sKey1).Split("|"c)(iCombo1 - 2)
                                                                        For Each state As String In prop.arlStates
                                                                            If sIntended = state.ToLower Then
                                                                                .sKey2 = state
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                    Else
                                                                        .sKey2 = prop.arlStates(iCombo1 - 2)
                                                                    End If
                                                                End If
                                                            Next
                                                        End If
                                                    Else
                                                        For Each prop As clsProperty In ob.htblActualProperties.Values
                                                            If prop.Key.IndexOf("|"c) > 0 Then
                                                                If dictDodgyStates.ContainsKey(.sKey1) Then
                                                                    Dim sIntended As String = dictDodgyStates(.sKey1).Split("|"c)(iCombo1)
                                                                    For Each state As String In prop.arlStates
                                                                        If sIntended = state.ToLower Then
                                                                            .sKey2 = state
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                Else
                                                                    .sKey2 = prop.arlStates(iCombo1)
                                                                End If
                                                            End If
                                                        Next
                                                    End If
                                                Case Else
                                                    Dim iOffset As Integer = 0
                                                    If .sKey1 = "ReferencedObject" OrElse ob.Openable Then
                                                        If .sKey1 = "ReferencedObject" OrElse ob.Lockable Then iOffset = 3 Else iOffset = 2
                                                    End If
                                                    For Each prop As clsProperty In ob.htblActualProperties.Values
                                                        If prop.Key.IndexOf("|"c) > 0 Then
                                                            .sKey2 = prop.arlStates(iCombo1 - iOffset)
                                                        End If
                                                    Next
                                            End Select


                                        Case 2 ' Task status
                                            .eType = clsRestriction.RestrictionTypeEnum.Task
                                            .sKey1 = "Task" & iCombo0 + iStartTask
                                            If iCombo1 = 0 Then
                                                .eMust = clsRestriction.MustEnum.Must
                                            Else
                                                .eMust = clsRestriction.MustEnum.MustNot
                                            End If
                                            .eTask = clsRestriction.TaskEnum.Complete

                                        Case 3 ' Characters
                                            .eType = clsRestriction.RestrictionTypeEnum.Character
                                            Select Case iCombo0
                                                Case 0
                                                    .sKey1 = "%Player%"
                                                Case 1
                                                    .sKey1 = "ReferencedCharacter"
                                                Case Else
                                                    .sKey1 = "Character" & iCombo0 - 1 + iStartChar
                                            End Select
                                            Select Case iCombo1
                                                Case 0 ' Same room as
                                                    .eMust = clsRestriction.MustEnum.Must
                                                    .eCharacter = clsRestriction.CharacterEnum.BeInSameLocationAsCharacter
                                                Case 1 ' Not same room as
                                                    .eMust = clsRestriction.MustEnum.MustNot
                                                    .eCharacter = clsRestriction.CharacterEnum.BeInSameLocationAsCharacter
                                                Case 2 ' Alone
                                                    .eMust = clsRestriction.MustEnum.Must
                                                    .eCharacter = clsRestriction.CharacterEnum.BeAlone
                                                Case 3 ' Not alone
                                                    .eMust = clsRestriction.MustEnum.MustNot
                                                    .eCharacter = clsRestriction.CharacterEnum.BeAlone
                                                Case 4 ' standing on 
                                                    .eMust = clsRestriction.MustEnum.Must
                                                    .eCharacter = clsRestriction.CharacterEnum.BeStandingOnObject
                                                Case 5 ' sitting on 
                                                    .eMust = clsRestriction.MustEnum.Must
                                                    .eCharacter = clsRestriction.CharacterEnum.BeSittingOnObject
                                                Case 6 ' lying on
                                                    .eMust = clsRestriction.MustEnum.Must
                                                    .eCharacter = clsRestriction.CharacterEnum.BeLyingOnObject
                                                Case 7  ' gender
                                                    .eMust = clsRestriction.MustEnum.Must
                                                    .eCharacter = clsRestriction.CharacterEnum.BeOfGender
                                            End Select
                                            Select Case iCombo1
                                                Case 0, 1
                                                    Select Case iCombo2
                                                        Case 0
                                                            .sKey2 = "%Player%"
                                                        Case 1
                                                            .sKey2 = "ReferencedCharacter"
                                                        Case Else
                                                            .sKey2 = "Character" & iCombo2 - 1 + iStartChar
                                                    End Select
                                                Case 4
                                                    ' Standables
                                                    Select Case iCombo2
                                                        Case 0
                                                            ' The floor
                                                            .sKey2 = "TheFloor"
                                                        Case Else
                                                            .sKey2 = GetObKey(iCombo2 - 1, ComboEnum.Standable)
                                                    End Select
                                                Case 5
                                                    ' Sittables
                                                    Select Case iCombo2
                                                        Case 0
                                                            ' The floor
                                                            .sKey2 = "TheFloor"
                                                        Case Else
                                                            .sKey2 = GetObKey(iCombo2 - 1, ComboEnum.Sittable)
                                                    End Select
                                                Case 6
                                                    ' Lyables
                                                    Select Case iCombo2
                                                        Case 0
                                                            ' The floor
                                                            .sKey2 = "TheFloor"
                                                        Case Else
                                                            .sKey2 = GetObKey(iCombo2 - 1, ComboEnum.Lieable)
                                                    End Select

                                                Case 7
                                                    ' Gender
                                                    .sKey2 = CType(iCombo2, clsCharacter.GenderEnum).ToString
                                            End Select
                                        Case 4 ' Variables
                                            .eType = clsRestriction.RestrictionTypeEnum.Variable
                                            Select Case iCombo0
                                                Case 0
                                                    .sKey1 = "ReferencedNumber"
                                                Case 1
                                                    .sKey1 = "ReferencedText"
                                                Case Else
                                                    .sKey1 = "Variable" & (iCombo0 - 1)
                                            End Select
                                            .sKey2 = "" ' Arrays not used in v4
                                            .eMust = clsRestriction.MustEnum.Must
                                            Select Case iCombo1
                                                Case 0, 10
                                                    .eVariable = clsRestriction.VariableEnum.LessThan
                                                Case 1, 11
                                                    .eVariable = clsRestriction.VariableEnum.LessThanOrEqualTo
                                                Case 2, 12
                                                    .eVariable = clsRestriction.VariableEnum.EqualTo
                                                Case 3, 13
                                                    .eVariable = clsRestriction.VariableEnum.GreaterThanOrEqualTo
                                                Case 4, 14
                                                    .eVariable = clsRestriction.VariableEnum.GreaterThan
                                                Case 5, 15
                                                    .eVariable = clsRestriction.VariableEnum.EqualTo
                                                    .eMust = clsRestriction.MustEnum.MustNot
                                            End Select
                                            If iCombo1 < 10 Then
                                                .IntValue = iCombo2
                                                If v >= 4 Then .StringValue = sText
                                            Else
                                                .IntValue = Integer.MinValue
                                                .StringValue = "Variable" & iCombo2
                                            End If
                                            'GetLine(bAdventure, iPos) ' combo(2)
                                            'GetLine(bAdventure, iPos) ' text
                                    End Select

                                    .oMessage = New Description(ConvText(GetLine(bAdventure, iPos)))
                                End With
                                .arlRestrictions.Add(NewRestriction)
                            Next

                            Dim iNumActions As Integer = CInt(GetLine(bAdventure, iPos))
                            For i As Integer = 1 To iNumActions
                                Dim NewAction As New clsAction
                                With NewAction
                                    Dim m As Integer = CInt(GetLine(bAdventure, iPos)) ' mode
                                    Dim iCombo0 As Integer = CInt(GetLine(bAdventure, iPos))
                                    Dim iCombo1, iCombo2, iCombo3 As Integer
                                    Dim sExpression As String = ""
                                    If v < 4 Then
                                        If m < 4 OrElse m = 6 Then iCombo1 = CInt(GetLine(bAdventure, iPos))
                                        If m = 0 OrElse m = 1 OrElse m = 3 OrElse m = 6 Then iCombo2 = CInt(GetLine(bAdventure, iPos))
                                        If m > 4 Then m += 1
                                        If m = 1 AndAlso iCombo1 = 2 Then iCombo2 += 2
                                        If m = 7 Then
                                            If iCombo0 >= 5 AndAlso iCombo0 <= 6 Then iCombo0 += 2
                                            If iCombo0 = 7 Then iCombo0 = 11
                                        End If
                                    Else
                                        If m < 4 Or m = 5 Or m = 6 Or m = 7 Then iCombo1 = CInt(GetLine(bAdventure, iPos))
                                        If m = 0 Or m = 1 Or m = 3 Or m = 6 Or m = 7 Then iCombo2 = CInt(GetLine(bAdventure, iPos))
                                    End If

                                    If m = 3 Then
                                        If v < 4 Then
                                            If iCombo1 = 5 Then
                                                sExpression = GetLine(bAdventure, iPos) ' expression
                                            Else
                                                iCombo3 = CInt(GetLine(bAdventure, iPos)) ' combo(3)
                                            End If
                                        Else
                                            sExpression = GetLine(bAdventure, iPos) ' expression
                                            iCombo3 = CInt(GetLine(bAdventure, iPos)) ' combo(3)
                                        End If
                                    End If
                                    Select Case m
                                        Case 0 ' Move object
                                            .eItem = clsAction.ItemEnum.MoveObject
                                            Select Case iCombo0
                                                Case 0
                                                    .eMoveObjectWhat = clsAction.MoveObjectWhatEnum.EverythingHeldBy
                                                    .sKey1 = THEPLAYER
                                                    '.sKey1 = "AllHeldObjects"
                                                Case 1
                                                    .eMoveObjectWhat = clsAction.MoveObjectWhatEnum.EverythingWornBy
                                                    .sKey1 = THEPLAYER
                                                    '.sKey1 = "AllWornObjects"
                                                Case 2
                                                    .eMoveObjectWhat = clsAction.MoveObjectWhatEnum.Object
                                                    .sKey1 = "ReferencedObject"
                                                Case Else
                                                    .eMoveObjectWhat = clsAction.MoveObjectWhatEnum.Object
                                                    .sKey1 = GetObKey(iCombo0 - 3, ComboEnum.Dynamic)
                                            End Select

                                            Select Case iCombo1
                                                Case 0
                                                    .eMoveObjectTo = clsAction.MoveObjectToEnum.ToLocation
                                                    If iCombo2 = 0 Then
                                                        .sKey2 = "Hidden"
                                                    Else
                                                        .sKey2 = "Location" & iCombo2 + iStartLocations
                                                    End If
                                                Case 1
                                                    .eMoveObjectTo = clsAction.MoveObjectToEnum.ToLocationGroup
                                                Case 2
                                                    .eMoveObjectTo = clsAction.MoveObjectToEnum.InsideObject
                                                    .sKey2 = GetObKey(iCombo2, ComboEnum.Container)
                                                Case 3
                                                    .eMoveObjectTo = clsAction.MoveObjectToEnum.OntoObject
                                                    .sKey2 = GetObKey(iCombo2, ComboEnum.Surface)
                                                Case 4
                                                    .eMoveObjectTo = clsAction.MoveObjectToEnum.ToCarriedBy
                                                Case 5
                                                    .eMoveObjectTo = clsAction.MoveObjectToEnum.ToWornBy
                                                Case 6
                                                    .eMoveObjectTo = clsAction.MoveObjectToEnum.ToSameLocationAs
                                            End Select

                                            If iCombo1 > 3 Then
                                                Select Case iCombo2
                                                    Case 0
                                                        .sKey2 = "%Player%"
                                                    Case 1
                                                        .sKey2 = "ReferencedCharacter"
                                                    Case Else
                                                        .sKey2 = "Character" & iCombo2 - 1 + iStartChar
                                                End Select
                                            End If

                                        Case 1 ' Move character
                                            .eItem = clsAction.ItemEnum.MoveCharacter
                                            .eMoveCharacterWho = clsAction.MoveCharacterWhoEnum.Character

                                            Select Case iCombo0
                                                Case 0
                                                    .sKey1 = THEPLAYER
                                                Case 1
                                                    .sKey1 = "ReferencedCharacter"
                                                Case Else
                                                    .sKey1 = "Character" & iCombo0 - 1 + iStartChar
                                            End Select

                                            Select Case iCombo1
                                                Case 0
                                                    .eMoveCharacterTo = clsAction.MoveCharacterToEnum.ToLocation
                                                    If .sKey1 = "%Player%" Then
                                                        .sKey2 = "Location" & iCombo2 + iStartLocations + 1
                                                    Else
                                                        .sKey2 = "Location" & iCombo2 + iStartLocations
                                                    End If
                                                    If .sKey2 = "Location0" Then .sKey2 = "Hidden"
                                                Case 1
                                                    .eMoveCharacterTo = clsAction.MoveCharacterToEnum.ToLocationGroup

                                                Case 2
                                                    .eMoveCharacterTo = clsAction.MoveCharacterToEnum.ToSameLocationAs
                                                    Select Case iCombo2
                                                        Case 0
                                                            .sKey2 = "%Player%"
                                                        Case 1
                                                            .sKey2 = "ReferencedCharacter"
                                                        Case Else
                                                            .sKey2 = "Character" & iCombo2 - 2 + iStartChar
                                                    End Select
                                                Case 3
                                                    .eMoveCharacterTo = clsAction.MoveCharacterToEnum.ToStandingOn
                                                    Select Case iCombo2
                                                        Case 0
                                                            ' The floor
                                                            .sKey2 = THEFLOOR
                                                        Case Else
                                                            .sKey2 = GetObKey(iCombo2 - 1, ComboEnum.Standable)
                                                    End Select
                                                Case 4
                                                    .eMoveCharacterTo = clsAction.MoveCharacterToEnum.ToSittingOn
                                                    Select Case iCombo2
                                                        Case 0
                                                            ' The floor
                                                            .sKey2 = THEFLOOR
                                                        Case Else
                                                            .sKey2 = GetObKey(iCombo2 - 1, ComboEnum.Sittable)
                                                    End Select
                                                Case 5
                                                    .eMoveCharacterTo = clsAction.MoveCharacterToEnum.ToLyingOn
                                                    Select Case iCombo2
                                                        Case 0
                                                            ' The floor
                                                            .sKey2 = THEFLOOR
                                                        Case Else
                                                            .sKey2 = GetObKey(iCombo2 - 1, ComboEnum.Lieable)
                                                    End Select
                                            End Select

                                        Case 2 ' Change ob status
                                            .eItem = clsAction.ItemEnum.SetProperties
                                            .sKey1 = GetObKey(iCombo0, ComboEnum.WithStateOrOpenable)
                                            Dim ob As clsObject = Adventure.htblObjects(.sKey1)
                                            Select Case iCombo1
                                                Case 0
                                                    If ob.Openable Then
                                                        .sKey2 = "OpenStatus"
                                                        .sPropertyValue = "Open"
                                                    Else
                                                        For Each prop As clsProperty In ob.htblActualProperties.Values
                                                            If prop.Key.IndexOf("|"c) > 0 Then
                                                                .sKey2 = prop.Key
                                                                .sPropertyValue = prop.arlStates(iCombo1)
                                                            End If
                                                        Next
                                                    End If
                                                Case 1
                                                    If ob.Openable Then
                                                        .sKey2 = "OpenStatus"
                                                        .sPropertyValue = "Closed"
                                                    Else
                                                        For Each prop As clsProperty In ob.htblActualProperties.Values
                                                            If prop.Key.IndexOf("|"c) > 0 Then
                                                                .sKey2 = prop.Key
                                                                .sPropertyValue = prop.arlStates(iCombo1)
                                                            End If
                                                        Next
                                                    End If
                                                Case 2
                                                    If ob.Openable Then
                                                        If ob.Lockable Then
                                                            .sKey2 = "OpenStatus"
                                                            .sPropertyValue = "Locked"
                                                        Else
                                                            For Each prop As clsProperty In ob.htblActualProperties.Values
                                                                If prop.Key.IndexOf("|"c) > 0 Then
                                                                    .sKey2 = prop.Key
                                                                    .sPropertyValue = prop.arlStates(iCombo1 - 2)
                                                                End If
                                                            Next
                                                        End If
                                                    Else
                                                        For Each prop As clsProperty In ob.htblActualProperties.Values
                                                            If prop.Key.IndexOf("|"c) > 0 Then
                                                                .sKey2 = prop.Key
                                                                .sPropertyValue = prop.arlStates(iCombo1)
                                                            End If
                                                        Next
                                                    End If
                                                Case Else
                                                    Dim iOffset As Integer = 0
                                                    If ob.Openable Then
                                                        If ob.Lockable Then iOffset = 3 Else iOffset = 2
                                                    End If
                                                    For Each prop As clsProperty In ob.htblActualProperties.Values
                                                        If prop.Key.IndexOf("|"c) > 0 Then
                                                            .sKey2 = prop.Key
                                                            .sPropertyValue = prop.arlStates(iCombo1 - iOffset)
                                                        End If
                                                    Next
                                            End Select

                                        Case 3 ' Change variable
                                            .eItem = clsAction.ItemEnum.SetVariable
                                            .sKey1 = "Variable" & iCombo0 + 1 '+ iStartVariable
                                            .eVariables = clsAction.VariablesEnum.Assignment
                                            .StringValue = iCombo1 & Chr(1) & iCombo2 & Chr(1) & iCombo3 & Chr(1) & sExpression

                                        Case 4 ' Change score
                                            .eItem = clsAction.ItemEnum.SetVariable
                                            .sKey1 = "Score"
                                            .eVariables = clsAction.VariablesEnum.Assignment
                                            .StringValue = "1" & Chr(1) & iCombo0 & Chr(1) & "0" & Chr(1)
                                            '.IntValue = iCombo0

                                        Case 5 ' Set Task
                                            .eItem = clsAction.ItemEnum.SetTasks
                                            If iCombo0 = 0 Then
                                                .eSetTasks = clsAction.SetTasksEnum.Execute
                                            Else
                                                .eSetTasks = clsAction.SetTasksEnum.Unset
                                            End If
                                            .sKey1 = "Task" & iCombo1 + iStartTask + 1
                                            .StringValue = ""

                                        Case 6 ' End game
                                            .eItem = clsAction.ItemEnum.EndGame
                                            Select Case iCombo0
                                                Case 0
                                                    .eEndgame = clsAction.EndGameEnum.Win
                                                Case 1
                                                    .eEndgame = clsAction.EndGameEnum.Neutral
                                                Case 2, 3
                                                    .eEndgame = clsAction.EndGameEnum.Lose
                                            End Select
                                        Case 7 ' Battles
                                            ' TODO

                                    End Select
                                End With
                                .arlActions.Add(NewAction)
                            Next
                        End If

                        Dim sBrackSeq As String = ""
                        If v < 4 Then
                            If .arlRestrictions.Count > 0 Then sBrackSeq = "#"
                            For i As Integer = 1 To .arlRestrictions.Count - 1
                                sBrackSeq &= "A#"
                            Next
                            .arlRestrictions.BracketSequence = sBrackSeq
                        Else
                            sBrackSeq = GetLine(bAdventure, iPos)
                            If sBrackSeq <> "" AndAlso .arlRestrictions.BracketSequence <> "" Then
                                .arlRestrictions.BracketSequence &= "A"
                            End If
                            .arlRestrictions.BracketSequence &= sBrackSeq
                        End If
                        If .arlRestrictions.BracketSequence IsNot Nothing Then .arlRestrictions.BracketSequence = .arlRestrictions.BracketSequence.Replace("[", "((").Replace("]", "))")
                        If v >= 3.9 Then
                            If bSound Then
                                sFilename = GetLine(bAdventure, iPos) ' Filename
                                Dim sLoop As String = ""
                                If sFilename.EndsWith("##") Then
                                    sLoop = " loop=Y"
                                    sFilename = sFilename.Substring(0, sFilename.Length - 2)
                                End If
                                If sFilename <> "" Then .CompletionMessage(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & .CompletionMessage(0).Description
                                If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                            End If
                            If bGraphics Then
                                sFilename = GetLine(bAdventure, iPos) ' Filename
                                If sFilename <> "" Then .CompletionMessage(0).Description = "<img src=""" & sFilename & """>" & .CompletionMessage(0).Description
                                If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, True))
                            End If
                        End If
                    End With
                    .htblTasks.Add(NewTask, NewTask.Key)
                Next


                '----------------------------------------------------------------------------------
                ' Events
                '----------------------------------------------------------------------------------

                Dim iNumEvents As Integer = CInt(GetLine(bAdventure, iPos))
                For iEvent As Integer = 1 To iNumEvents
                    Dim NewEvent As New clsEvent
                    Dim sLocationKey As String = ""
                    With NewEvent
                        .Key = "Event" & iEvent.ToString
                        .Description = GetLine(bAdventure, iPos)
                        .WhenStart = CType(GetLine(bAdventure, iPos), clsEvent.WhenStartEnum)
                        If .WhenStart = clsEvent.WhenStartEnum.BetweenXandYTurns Then
                            .StartDelay.iFrom = CInt(GetLine(bAdventure, iPos)) - 1 ' Start1
                            .StartDelay.iTo = CInt(GetLine(bAdventure, iPos)) - 1 ' Start2
                        End If

                        If .WhenStart = clsEvent.WhenStartEnum.AfterATask Then
                            Dim sStartTask As String = "Task" & GetLine(bAdventure, iPos)
                            Dim ec As New EventOrWalkControl
                            ec.eControl = EventOrWalkControl.ControlEnum.Start
                            ec.sTaskKey = sStartTask
                            ReDim Preserve .EventControls(.EventControls.Length)
                            .EventControls(.EventControls.Length - 1) = ec
                        End If

                        .Repeating = CBool(GetLine(bAdventure, iPos))
                        Dim iTaskMode As Integer = CInt(GetLine(bAdventure, iPos)) ' task mode
                        .Length.iFrom = CInt(GetLine(bAdventure, iPos))
                        .Length.iTo = CInt(GetLine(bAdventure, iPos))
                        If .WhenStart = clsEvent.WhenStartEnum.BetweenXandYTurns Then
                            .Length.iFrom -= 1
                            .Length.iTo -= 1
                        End If
                        sBuffer = CStr(GetLine(bAdventure, iPos)) ' des1
                        If sBuffer <> "" Then
                            Dim se As New clsEvent.SubEvent(.Key)
                            se.eWhat = clsEvent.SubEvent.WhatEnum.DisplayMessage
                            se.eWhen = clsEvent.SubEvent.WhenEnum.FromStartOfEvent
                            se.ftTurns.iFrom = 0
                            se.ftTurns.iTo = 0
                            se.oDescription = New Description(ConvText(sBuffer))
                            ReDim Preserve .SubEvents(.SubEvents.Length)
                            .SubEvents(.SubEvents.Length - 1) = se
                        End If
                        sBuffer = CStr(GetLine(bAdventure, iPos)) ' des2
                        If sBuffer <> "" Then
                            Dim se As New clsEvent.SubEvent(.Key)
                            se.eWhat = clsEvent.SubEvent.WhatEnum.SetLook
                            se.eWhen = clsEvent.SubEvent.WhenEnum.FromStartOfEvent
                            se.ftTurns.iFrom = 0
                            se.ftTurns.iTo = 0
                            se.oDescription = New Description(ConvText(sBuffer))
                            ReDim Preserve .SubEvents(.SubEvents.Length)
                            .SubEvents(.SubEvents.Length - 1) = se
                        End If
                        Dim sEndMessage As String = CStr(GetLine(bAdventure, iPos)) ' des3                        

                        Dim iWhichRooms As Integer = CInt(GetLine(bAdventure, iPos))
                        Select Case iWhichRooms
                            Case 0 ' No rooms
                                sLocationKey = ""
                            Case 1 ' Single Room
                                sLocationKey = "Location" & (CInt(GetLine(bAdventure, iPos)) + 1)
                            Case 2 ' Multiple Rooms
                                Dim bShowRoom As Boolean
                                Dim arlShowInRooms As New StringArrayList
                                For n As Integer = 1 To iNumLocations 'Adventure.htblLocations.Count
                                    bShowRoom = CBool(GetLine(bAdventure, iPos))
                                    If bShowRoom Then arlShowInRooms.Add("Location" & n)
                                Next
                                sLocationKey = GetRoomGroupFromList(iPos, arlShowInRooms, "event '" & .Description & "'").Key
                            Case 3 ' All Rooms
                                sLocationKey = ALLROOMS
                        End Select

                        For i As Integer = 0 To 1
                            Dim iTask As Integer = CInt(GetLine(bAdventure, iPos))
                            Dim iCompleteOrNot As Integer = CInt(GetLine(bAdventure, iPos))
                            If iTask > 0 Then
                                Dim ec As New EventOrWalkControl
                                ec.eControl = CType(IIf(i = 0, EventOrWalkControl.ControlEnum.Suspend, EventOrWalkControl.ControlEnum.Resume), EventOrWalkControl.ControlEnum)
                                ec.sTaskKey = "Task" & (iTask - 1)
                                ec.eCompleteOrNot = CType(IIf(iCompleteOrNot = 0, EventOrWalkControl.CompleteOrNotEnum.Completion, EventOrWalkControl.CompleteOrNotEnum.UnCompletion), EventOrWalkControl.CompleteOrNotEnum)
                                ReDim Preserve .EventControls(.EventControls.Length)
                                .EventControls(.EventControls.Length - 1) = ec
                            End If
                            Dim iFrom As Integer = CInt(GetLine(bAdventure, iPos)) ' from(i)
                            sBuffer = CStr(GetLine(bAdventure, iPos)) ' ftext(i)
                            If sBuffer <> "" Then
                                Dim se As New clsEvent.SubEvent(.Key)
                                se.eWhat = clsEvent.SubEvent.WhatEnum.DisplayMessage
                                se.eWhen = clsEvent.SubEvent.WhenEnum.BeforeEndOfEvent
                                se.ftTurns.iFrom = iFrom
                                se.ftTurns.iTo = iFrom
                                se.oDescription = New Description(ConvText(sBuffer))
                                ReDim Preserve .SubEvents(.SubEvents.Length)
                                .SubEvents(.SubEvents.Length - 1) = se
                            End If
                        Next
                        If sEndMessage <> "" Then
                            Dim se As New clsEvent.SubEvent(.Key)
                            se.eWhat = clsEvent.SubEvent.WhatEnum.DisplayMessage
                            se.eWhen = clsEvent.SubEvent.WhenEnum.BeforeEndOfEvent
                            se.ftTurns.iFrom = 0
                            se.ftTurns.iTo = 0
                            se.oDescription = New Description(ConvText(sEndMessage))
                            ReDim Preserve .SubEvents(.SubEvents.Length)
                            .SubEvents(.SubEvents.Length - 1) = se
                        End If
                        Dim tas As clsTask = Nothing
                        Dim iDoneTask(1) As Boolean
                        Dim iMoveObs(2, 1) As Integer
                        For Each i As Integer In New Integer() {1, 2, 0}
                            For j As Integer = 0 To 1
                                iMoveObs(i, j) = CInt(GetLine(bAdventure, iPos))
                            Next
                        Next
                        For i As Integer = 0 To 2
                            Dim iObKey As Integer = iMoveObs(i, 0)
                            Dim iMoveTo As Integer = iMoveObs(i, 1)
                            If iObKey > 0 Then
                                Dim bNewTask As Boolean = True
                                If i = 1 AndAlso NewEvent.Length.iTo = 0 AndAlso iDoneTask(0) Then bNewTask = False
                                If i = 2 AndAlso ((NewEvent.Length.iTo = 0 AndAlso Not iDoneTask(1)) OrElse iDoneTask(1)) Then bNewTask = False

                                If bNewTask Then
                                    Dim bMultiple As Boolean = False
                                    If tas IsNot Nothing Then
                                        bMultiple = True
                                        tas.Description = "Generated task #" & i & " for event " & NewEvent.Description
                                    End If
                                    tas = New clsTask
                                    tas.Key = "GenTask" & (Adventure.htblTasks.Count + 1)
                                    tas.Description = "Generated task" & IIf(bMultiple, " #" & i + 1, "").ToString & " for event " & NewEvent.Description
                                    tas.Priority = iStartMaxPriority + Adventure.htblTasks.Count + 1
                                End If
                                If i < 2 Then iDoneTask(i) = True

                                With tas
                                    .TaskType = clsTask.TaskTypeEnum.System
                                    .Repeatable = True
                                    Dim act As New clsAction
                                    act.eItem = clsAction.ItemEnum.MoveObject
                                    act.sKey1 = "Object" & iObKey
                                    Select Case iMoveTo
                                        Case 0 ' Hidden
                                            act.eMoveObjectTo = clsAction.MoveObjectToEnum.ToLocation
                                            act.sKey2 = "Hidden"
                                        Case 1 ' Players hands
                                            If Adventure.htblObjects("Object" & iObKey).IsStatic Then
                                                act.eMoveObjectTo = clsAction.MoveObjectToEnum.ToLocation '  Don't allow for static
                                                act.sKey2 = "Hidden"
                                            Else
                                                act.eMoveObjectTo = clsAction.MoveObjectToEnum.ToCarriedBy
                                                act.sKey2 = "%Player%"
                                            End If
                                        Case 2 ' Same room as player
                                            act.eMoveObjectTo = clsAction.MoveObjectToEnum.ToSameLocationAs
                                            act.sKey2 = "%Player%"
                                        Case Else ' Locations
                                            act.eMoveObjectTo = clsAction.MoveObjectToEnum.ToLocation
                                            act.sKey2 = "Location" & (iMoveTo - 2)
                                    End Select

                                    .arlActions.Add(act)
                                End With
                                If bNewTask Then
                                    Adventure.htblTasks.Add(tas, tas.Key)
                                    Dim se As New clsEvent.SubEvent(.Key)
                                    se.eWhat = clsEvent.SubEvent.WhatEnum.ExecuteTask
                                    If i = 0 Then se.eWhen = clsEvent.SubEvent.WhenEnum.FromStartOfEvent Else se.eWhen = clsEvent.SubEvent.WhenEnum.BeforeEndOfEvent
                                    se.ftTurns.iFrom = 0
                                    se.ftTurns.iTo = 0
                                    se.sKey = tas.Key
                                    ReDim Preserve .SubEvents(.SubEvents.Length)
                                    .SubEvents(.SubEvents.Length - 1) = se
                                End If
                            End If
                        Next
                        For Each se As clsEvent.SubEvent In .SubEvents
                            If se.eWhat = clsEvent.SubEvent.WhatEnum.DisplayMessage OrElse se.eWhat = clsEvent.SubEvent.WhatEnum.SetLook Then se.sKey = sLocationKey
                        Next
                        Dim sExecuteTask As String = "Task" & GetLine(bAdventure, iPos)
                        If sExecuteTask <> "Task0" Then
                            Dim se As New clsEvent.SubEvent(.Key)
                            If iTaskMode = 0 Then
                                se.eWhat = clsEvent.SubEvent.WhatEnum.ExecuteTask
                            Else
                                se.eWhat = clsEvent.SubEvent.WhatEnum.UnsetTask
                            End If
                            se.eWhen = clsEvent.SubEvent.WhenEnum.BeforeEndOfEvent
                            se.ftTurns.iFrom = 0
                            se.ftTurns.iTo = 0
                            se.sKey = sExecuteTask
                            ReDim Preserve .SubEvents(.SubEvents.Length)
                            .SubEvents(.SubEvents.Length - 1) = se
                        End If
                        If v >= 3.9 Then
                            For i As Integer = 0 To 4
                                If bSound Then
                                    sFilename = GetLine(bAdventure, iPos) ' Filename
                                    If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                    If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                                End If
                                If bGraphics Then
                                    sFilename = GetLine(bAdventure, iPos) ' Filename
                                    If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                    If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, True))
                                End If
                            Next
                        End If
                    End With
                    .htblEvents.Add(NewEvent, NewEvent.Key)
                Next


                '----------------------------------------------------------------------------------
                ' Characters
                '----------------------------------------------------------------------------------

                Dim iNumChars As Integer = CInt(GetLine(bAdventure, iPos))
                For iChar As Integer = 1 To iNumChars
                    Dim NewChar As New clsCharacter
                    With NewChar
                        Dim sKey As String = "Character" & iChar.ToString
                        If a.htblCharacters.ContainsKey(sKey) Then
                            While a.htblCharacters.ContainsKey(sKey)
                                sKey = IncrementKey(sKey)
                            End While
                        End If
                        .Key = sKey
                        .CharacterType = clsCharacter.CharacterTypeEnum.NonPlayer
                        .ProperName = GetLine(bAdventure, iPos)
                        .Prefix = GetLine(bAdventure, iPos)
                        ConvertPrefix(.Article, .Prefix)
                        If v < 4 Then
                            Dim sAlias As String = GetLine(bAdventure, iPos)
                            If sAlias <> "" Then .arlDescriptors.Add(sAlias)
                        Else
                            Dim iNumAliases As Integer = CInt(GetLine(bAdventure, iPos))
                            For i As Integer = 1 To iNumAliases
                                .arlDescriptors.Add(GetLine(bAdventure, iPos))
                            Next
                            If iNumAliases = 0 AndAlso .Prefix = "" Then .Article = ""
                        End If

                        .Description = New Description(ConvText(GetLine(bAdventure, iPos)))
                        .Known = True

                        Dim iCharLoc As Integer = CInt(GetLine(bAdventure, iPos))
                        If iCharLoc > 0 Then
                            .Location.ExistWhere = clsCharacterLocation.ExistsWhereEnum.AtLocation
                            .Location.Key = "Location" & iCharLoc
                        Else
                            .Location.ExistWhere = clsCharacterLocation.ExistsWhereEnum.Hidden
                        End If

                        Dim p As clsProperty

                        Dim sDesc2 As String = GetLine(bAdventure, iPos)
                        Dim sDescTask As String = GetLine(bAdventure, iPos)
                        If sDescTask <> "0" Then
                            Dim sd As New SingleDescription
                            Dim rest As New clsRestriction
                            rest.eType = clsRestriction.RestrictionTypeEnum.Task
                            rest.sKey1 = "Task" & sDescTask
                            rest.eMust = clsRestriction.MustEnum.Must
                            rest.eTask = clsRestriction.TaskEnum.Complete
                            sd.Restrictions.Add(rest)
                            sd.Description = sDesc2
                            sd.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartDescriptionWithThis
                            sd.Restrictions.BracketSequence = "#"
                            .Description.Add(sd)
                        End If
                        Dim iNumSubjects As Integer = CInt(GetLine(bAdventure, iPos))
                        For i As Integer = 1 To iNumSubjects
                            Dim Topic As New clsTopic
                            With Topic
                                .Key = "Topic" & i
                                .Keywords = GetLine(bAdventure, iPos)
                                .Summary = "Ask about " & .Keywords
                                .oConversation = New Description(ConvText(GetLine(bAdventure, iPos)))
                                .bAsk = True
                                Dim iTask As Integer = CInt(GetLine(bAdventure, iPos)) ' Rep Task
                                If iTask > 0 Then
                                    Dim sd As New SingleDescription
                                    Dim rest As New clsRestriction
                                    rest.eType = clsRestriction.RestrictionTypeEnum.Task
                                    rest.sKey1 = "Task" & iTask
                                    rest.eMust = clsRestriction.MustEnum.Must
                                    rest.eTask = clsRestriction.TaskEnum.Complete
                                    sd.Restrictions.Add(rest)
                                    sd.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartDescriptionWithThis
                                    sd.Restrictions.BracketSequence = "#"
                                    .oConversation.Add(sd)
                                End If
                                .oConversation.Item(.oConversation.Count - 1).Description &= GetLine(bAdventure, iPos) ' Response 2
                            End With
                            .htblTopics.Add(Topic)
                        Next

                        Dim iNumWalks As Integer = CInt(GetLine(bAdventure, iPos))
                        Dim arlNewDescriptions As New StringArrayList
                        For i As Integer = 1 To iNumWalks
                            Dim Walk As New clsWalk
                            With Walk
                                .sKey = sKey
                                Dim sStartTaskKey As String = ""
                                Dim iNumberOfSteps As Integer = CInt(GetLine(bAdventure, iPos))
                                .Loops = CBool(GetLine(bAdventure, iPos))
                                Dim iTask As Integer = CInt(GetLine(bAdventure, iPos))
                                If iTask = 0 Then
                                    .StartActive = True
                                Else
                                    .StartActive = False
                                    sStartTaskKey = "Task" & iTask
                                    Dim wc As New EventOrWalkControl
                                    wc.eControl = EventOrWalkControl.ControlEnum.Start
                                    wc.sTaskKey = sStartTaskKey
                                    ReDim Preserve .WalkControls(.WalkControls.Length)
                                    .WalkControls(.WalkControls.Length - 1) = wc
                                End If
                                Dim iCharTask As Integer = CInt(GetLine(bAdventure, iPos)) ' Runtask
                                Dim iObFind As Integer = CInt(GetLine(bAdventure, iPos)) ' Obfind
                                Dim iObTask As Integer = CInt(GetLine(bAdventure, iPos)) ' Obtask
                                If iObFind > 0 AndAlso iObTask > 0 Then
                                    Dim sw As New clsWalk.SubWalk
                                    sw.eWhat = clsWalk.SubWalk.WhatEnum.ExecuteTask
                                    sw.eWhen = clsWalk.SubWalk.WhenEnum.ComesAcross
                                    sw.sKey = "Object" & iObFind
                                    sw.sKey2 = "Task" & iObTask
                                    ReDim Preserve .SubWalks(.SubWalks.Length)
                                    .SubWalks(.SubWalks.Length - 1) = sw
                                End If
                                Dim sTerminateTaskKey As String = "Task" & GetLine(bAdventure, iPos)
                                If sTerminateTaskKey <> "Task0" Then
                                    Dim wc As New EventOrWalkControl
                                    wc.eControl = EventOrWalkControl.ControlEnum.Stop
                                    wc.sTaskKey = sTerminateTaskKey
                                    ReDim Preserve .WalkControls(.WalkControls.Length)
                                    .WalkControls(.WalkControls.Length - 1) = wc
                                End If
                                If v >= 3.9 Then
                                    Dim iCharFind As Integer = 0
                                    If v >= 4 Then iCharFind = SafeInt(GetLine(bAdventure, iPos)) ' Who
                                    If iCharTask > 0 Then
                                        Dim sw As New clsWalk.SubWalk
                                        sw.eWhat = clsWalk.SubWalk.WhatEnum.ExecuteTask
                                        sw.eWhen = clsWalk.SubWalk.WhenEnum.ComesAcross
                                        If v >= 4 AndAlso iCharFind > 0 Then
                                            sw.sKey = "Character" & iCharFind
                                        Else
                                            sw.sKey = "%Player%"
                                        End If
                                        sw.sKey2 = "Task" & iCharTask
                                        ReDim Preserve .SubWalks(.SubWalks.Length)
                                        .SubWalks(.SubWalks.Length - 1) = sw
                                    End If
                                    Dim sNewDescription As String = GetLine(bAdventure, iPos)
                                    If sNewDescription <> "" Then
                                        arlNewDescriptions.Add(sStartTaskKey)
                                        arlNewDescriptions.Add(sNewDescription)
                                    End If
                                End If

                                'Dim sDescription As String = ""
                                For j As Integer = 1 To iNumberOfSteps
                                    Dim Stp As New clsWalk.clsStep
                                    With Stp
                                        Dim iLocation As Integer = CInt(GetLine(bAdventure, iPos))
                                        Select Case iLocation
                                            Case 0 ' Hidden
                                                .sLocation = "Hidden"
                                            Case 1 ' Follow Player
                                                .sLocation = "%Player%"
                                            Case Else ' Locations
                                                If iLocation - 1 > Adventure.htblLocations.Count Then
                                                    ' Location Group
                                                    .sLocation = "Group" & iLocation - Adventure.htblLocations.Count - 1
                                                Else
                                                    ' Location
                                                    .sLocation = "Location" & iLocation - 1
                                                End If
                                        End Select
                                        Dim iWaitTurns As Integer = CInt(GetLine(bAdventure, iPos))
                                        .ftTurns.iFrom = iWaitTurns
                                        .ftTurns.iTo = iWaitTurns
                                    End With
                                    .arlSteps.Add(Stp)
                                Next
                                .Description = .GetDefaultDescription
                            End With
                            .arlWalks.Add(Walk)
                        Next
                        Dim bShowMove As Boolean = CBool(GetLine(bAdventure, iPos))
                        Dim sFromDesc As String = Nothing
                        Dim sToDesc As String = Nothing
                        If bShowMove Then
                            p = Adventure.htblAllProperties("ShowEnterExit").Copy
                            p.Selected = True
                            .AddProperty(p)
                            sFromDesc = GetLine(bAdventure, iPos)
                            p = Adventure.htblAllProperties("CharEnters").Copy
                            p.Selected = True
                            p.StringData = New Description(ConvText(sFromDesc))
                            .AddProperty(p)
                            sToDesc = GetLine(bAdventure, iPos)
                            p = Adventure.htblAllProperties("CharExits").Copy
                            p.Selected = True
                            p.StringData = New Description(ConvText(sToDesc))
                            .AddProperty(p)
                        End If
                        Dim sIsHereDesc As String = GetLine(bAdventure, iPos)
                        If sIsHereDesc = "#" Then sIsHereDesc = "%CharacterName[" & sKey & "]% is here."
                        If sIsHereDesc <> "" OrElse arlNewDescriptions.Count > 0 Then
                            p = Adventure.htblAllProperties("CharHereDesc").Copy
                            p.Selected = True
                            .AddProperty(p)
                        End If
                        If sIsHereDesc <> "" Then .SetPropertyValue("CharHereDesc", New Description(ConvText(sIsHereDesc)))
                        For i As Integer = 0 To arlNewDescriptions.Count - 1 Step 2
                            Dim sd As New SingleDescription
                            Dim rest As New clsRestriction
                            rest.eType = clsRestriction.RestrictionTypeEnum.Task
                            rest.sKey1 = arlNewDescriptions(i)
                            rest.eMust = clsRestriction.MustEnum.Must
                            rest.eTask = clsRestriction.TaskEnum.Complete
                            sd.Restrictions.Add(rest)
                            sd.Restrictions.BracketSequence = "#"
                            sd.Description = arlNewDescriptions(i + 1)
                            sd.eDisplayWhen = SingleDescription.DisplayWhenEnum.StartDescriptionWithThis
                            .GetProperty("CharHereDesc").StringData.Add(sd)
                        Next
                        If v >= 3.9 Then
                            .Gender = CType(GetLine(bAdventure, iPos), clsCharacter.GenderEnum)
                            For i As Integer = 0 To 3
                                If bSound Then
                                    sFilename = GetLine(bAdventure, iPos) ' Filename
                                    If sFilename <> "" Then
                                        Dim sLoop As String = ""
                                        If sFilename.EndsWith("##") Then
                                            sLoop = " loop=Y"
                                            sFilename = sFilename.Substring(0, sFilename.Length - 2)
                                        End If
                                        If i = 0 Then .Description(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & .Description(0).Description
                                        If i = 1 Then .Description(1).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & .Description(1).Description
                                        If i = 2 Then .GetProperty("CharEnters").StringData(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & .GetProperty("CharEnters").StringData(0).Description
                                        If i = 3 Then .GetProperty("CharExits").StringData(0).Description = "<audio play src=""" & sFilename & """" & sLoop & ">" & .GetProperty("CharExits").StringData(0).Description
                                    End If
                                    If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                    If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, False))
                                End If
                                If bGraphics Then
                                    sFilename = GetLine(bAdventure, iPos) ' Filename
                                    If sFilename <> "" Then
                                        If i = 0 Then .Description(0).Description = "<img src=""" & sFilename & """>" & .Description(0).Description
                                        If i = 1 Then .Description(1).Description = "<img src=""" & sFilename & """>" & .Description(1).Description
                                        If i = 2 Then .GetProperty("CharEnters").StringData(0).Description = "<img src=""" & sFilename & """>" & .GetProperty("CharEnters").StringData(0).Description
                                        If i = 3 Then .GetProperty("CharExits").StringData(0).Description = "<img src=""" & sFilename & """>" & .GetProperty("CharExits").StringData(0).Description
                                    End If
                                    If v >= 4 Then iFilesize = SafeInt(GetLine(bAdventure, iPos)) ' Filesize
                                    If sFilename <> "" AndAlso iFilesize > 0 Then a.dictv4Media.Add(sFilename, New clsAdventure.v4Media(0, iFilesize, True))
                                End If
                            Next
                            If bBattleSystem Then
                                GetLine(bAdventure, iPos) ' Attitude
                                GetLine(bAdventure, iPos) ' Min Stamina
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Max Stamina
                                GetLine(bAdventure, iPos) ' Min Strength
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Max Strength
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Min Accuracy
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Max Accuracy
                                GetLine(bAdventure, iPos) ' Min Defence
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Max Defence
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Min Agility
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Max Agility
                                GetLine(bAdventure, iPos) ' Speed
                                GetLine(bAdventure, iPos) ' Die Task
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Recovery
                                If v >= 4 Then GetLine(bAdventure, iPos) ' Low Task
                            End If
                        End If
                    End With
                    .htblCharacters.Add(NewChar, NewChar.Key)
                Next


                '----------------------------------------------------------------------------------
                ' Groups
                '----------------------------------------------------------------------------------

                ' Only room groups defined in 4.0 files
                Dim iNumGroups As Integer = CInt(GetLine(bAdventure, iPos))
                For iGroup As Integer = 1 To iNumGroups
                    Dim NewGroup As New clsGroup
                    With NewGroup
                        .Key = "Group" & iGroup.ToString
                        .Name = GetLine(bAdventure, iPos)
                        Dim bIncluded As Boolean
                        For i As Integer = 1 To iNumLocations
                            bIncluded = CBool(GetLine(bAdventure, iPos))
                            If bIncluded Then .arlMembers.Add("Location" & i.ToString)
                        Next
                    End With
                    .htblGroups.Add(NewGroup, NewGroup.Key)
                Next

                ' Sort out anything which needed groups defined
                For Each c As clsCharacter In Adventure.htblCharacters.Values
                    For Each w As clsWalk In c.arlWalks
                        For Each g As clsGroup In Adventure.htblGroups.Values
                            If w.Description.Contains("<" & g.Key & ">") Then
                                w.Description = w.Description.Replace("<" & g.Key & ">", g.Name)
                            End If
                        Next
                    Next
                Next


                '----------------------------------------------------------------------------------
                ' Synonyms
                '----------------------------------------------------------------------------------

                Dim iNumSyn As Integer = CInt(GetLine(bAdventure, iPos))
                For iSyn As Integer = 1 To iNumSyn
                    Dim sTo As String = GetLine(bAdventure, iPos) ' System Command
                    Dim sFrom As String = GetLine(bAdventure, iPos) ' Alternative Command
                    Dim synNew As clsSynonym = Nothing
                    For Each syn As clsSynonym In Adventure.htblSynonyms.Values
                        If syn.ChangeTo = sTo Then
                            synNew = syn
                            Exit For
                        End If
                    Next
                    If synNew Is Nothing Then
                        synNew = New clsSynonym
                        synNew.Key = "Synonym" & iSyn.ToString
                    End If
                    With synNew
                        .ChangeTo = sTo
                        .ChangeFrom.Add(sFrom)
                    End With
                    If Not .htblSynonyms.ContainsKey(synNew.Key) Then .htblSynonyms.Add(synNew)
                Next



                If v >= 3.9 Then

                    '----------------------------------------------------------------------------------
                    ' Variables
                    '----------------------------------------------------------------------------------

                    Dim iNumVariables As Integer = CInt(GetLine(bAdventure, iPos))
                    For iVar As Integer = 1 To iNumVariables
                        Dim NewVariable As New clsVariable
                        With NewVariable
                            .Key = "Variable" & iVar.ToString
                            .Name = GetLine(bAdventure, iPos)
                            If v < 4 Then
                                .Type = clsVariable.VariableTypeEnum.Numeric
                                .IntValue = CInt(GetLine(bAdventure, iPos))
                            Else
                                .Type = CType(GetLine(bAdventure, iPos), clsVariable.VariableTypeEnum)
                                If .Type = clsVariable.VariableTypeEnum.Numeric Then
                                    .IntValue = CInt(GetLine(bAdventure, iPos))
                                Else
                                    .StringValue = GetLine(bAdventure, iPos)
                                End If
                            End If

                        End With
                        .htblVariables.Add(NewVariable, NewVariable.Key)
                    Next

                    ' Change Variable names in Actions/Restrictions, and sort assignments
                    For Each tas As clsTask In Adventure.htblTasks.Values
                        For Each rest As clsRestriction In tas.arlRestrictions
                            If rest.eType = clsRestriction.RestrictionTypeEnum.Variable Then
                                If rest.sKey1 = "ReferencedText" OrElse Adventure.htblVariables.ContainsKey(rest.sKey1) Then
                                    If rest.sKey1 = "ReferencedText" OrElse Adventure.htblVariables(rest.sKey1).Type = clsVariable.VariableTypeEnum.Text Then
                                        Select Case rest.eVariable
                                            Case clsRestriction.VariableEnum.LessThan
                                                rest.eVariable = clsRestriction.VariableEnum.EqualTo
                                            Case clsRestriction.VariableEnum.LessThanOrEqualTo
                                                rest.eVariable = clsRestriction.VariableEnum.EqualTo
                                                rest.eMust = clsRestriction.MustEnum.MustNot
                                        End Select
                                        rest.StringValue = """" & rest.StringValue & """"
                                    End If
                                End If
                            End If
                        Next
                        For Each act As clsAction In tas.arlActions
                            If act.eItem = clsAction.ItemEnum.SetVariable Then
                                Dim iCombo1 As Integer = CInt(act.StringValue.Split(Chr(1))(0))
                                Dim iCombo2 As Integer = CInt(act.StringValue.Split(Chr(1))(1))
                                Dim iCombo3 As Integer = CInt(act.StringValue.Split(Chr(1))(2))
                                Dim sExpression As String = act.StringValue.Split(Chr(1))(3)
                                If a.htblVariables(act.sKey1).Type = clsVariable.VariableTypeEnum.Numeric Then
                                    Select Case iCombo1
                                        Case 0 ' to exact value
                                            act.StringValue = iCombo2.ToString
                                        Case 1 ' by exact value
                                            act.StringValue = "%" & a.htblVariables(act.sKey1).Name & "% + " & iCombo2.ToString
                                        Case 2 ' To Random value between X and Y
                                            act.StringValue = "Rand(" & iCombo2 & ", " & iCombo3 & ")"
                                        Case 3 ' By Random value between X and Y
                                            act.StringValue = "%" & a.htblVariables(act.sKey1).Name & "% + Rand(" & iCombo2 & ", " & iCombo3 & ")"
                                        Case 4 ' to referenced number
                                            act.StringValue = "%number%"
                                        Case 5 ' to expression
                                            act.StringValue = sExpression
                                        Case 6, 7, 8, 9, 10
                                            act.StringValue = ""
                                    End Select
                                Else
                                    Select Case iCombo1
                                        Case 0 ' exact text
                                            act.StringValue = """" & sExpression.Replace("""", "\""") & """"
                                        Case 1 ' to referenced text
                                            act.StringValue = "%text%"
                                        Case 2 ' to expression
                                            act.StringValue = sExpression
                                    End Select
                                End If

                            End If
                            If sInstr(act.StringValue, "Variable") > 0 Then
                                For iVar As Integer = Adventure.htblVariables.Count To 1 Step -1
                                    If Adventure.htblVariables.ContainsKey("Variable" & iVar) Then act.StringValue = act.StringValue.Replace("Variable" & iVar, "%" & Adventure.htblVariables("Variable" & iVar).Name & "%")
                                Next
                            End If
                        Next
                    Next



                    '----------------------------------------------------------------------------------
                    ' ALRs
                    '----------------------------------------------------------------------------------

                    Dim iNumALR As Integer = CInt(GetLine(bAdventure, iPos))
                    For iALR As Integer = 1 To iNumALR
                        Dim NewALR As New clsALR
                        With NewALR
                            .Key = "ALR" & iALR.ToString
                            .OldText = GetLine(bAdventure, iPos)
                            .NewText = New Description(ConvText(GetLine(bAdventure, iPos)))
                        End With
                        .htblALRs.Add(NewALR, NewALR.Key)
                    Next
                End If


                Dim bSetFont As Boolean = CBool(SafeInt(GetLine(bAdventure, iPos))) ' Set Font?
                If bSetFont Then
                    Dim sFont As String = GetLine(bAdventure, iPos) ' Font
                    If sFont.Contains(",") Then
                        Adventure.DefaultFontName = sFont.Split(","c)(0)
                        Adventure.DefaultFontSize = SafeInt(sFont.Split(","c)(1))
                    End If
                End If

                If v >= 4 Then
                    Dim iMediaOffset As Integer = bAdvZLib.Length + 23
                    For Each m As clsAdventure.v4Media In a.dictv4Media.Values
                        m.iOffset = iMediaOffset
                        iMediaOffset += m.iLength + 1
                    Next
                End If

                '--- the rest ---
                While iPos < bAdventure.Length - 1
                    Dim s3 As String = GetLine(bAdventure, iPos)
                End While

                ' Make sure all the 'seen's are set
                For Each ch As clsCharacter In Adventure.htblCharacters.Values
                    ch.Move(ch.Location)
                Next
                .Map = New clsMap
                .Map.RecalculateLayout()
            End With

        Catch ex As Exception
            ErrMsg("Error loading Adventure", ex)
            Return False
        Finally
            bAdventure = Nothing
        End Try

        Return True

    End Function

#If Not Adravalon Then
    Public Function Getv4Image(ByVal sFilename As String) As Drawing.Image

        If Adventure.dictv4Media.ContainsKey(sFilename) Then
            With Adventure.dictv4Media(sFilename)
                Dim stmFile As New IO.FileStream(Adventure.FullPath, IO.FileMode.Open, IO.FileAccess.Read)
                stmFile.Position = .iOffset
                Dim bytMedia(.iLength - 1) As Byte
                stmFile.Read(bytMedia, 0, .iLength)
                stmFile.Close()

                Dim msImage As New IO.MemoryStream(bytMedia)
                Return New Bitmap(msImage)
            End With
        Else
            ErrMsg("File " & sFilename & " not found in index.")
        End If
        Return Nothing

    End Function


    Public Function Getv4Audio(ByVal sFilename As String, ByVal sOutputFile As String) As Boolean

        If Adventure.dictv4Media.ContainsKey(sFilename) Then
            With Adventure.dictv4Media(sFilename)
                Dim stmFile As New IO.FileStream(Adventure.FullPath, IO.FileMode.Open, IO.FileAccess.Read)
                stmFile.Position = .iOffset
                Dim bytMedia(.iLength - 1) As Byte
                stmFile.Read(bytMedia, 0, .iLength)
                stmFile.Close()

                If sOutputFile <> "" Then
                    Dim stmOutput As New IO.FileStream(sOutputFile, IO.FileMode.Create)
                    stmOutput.Write(bytMedia, 0, bytMedia.Length - 1)
                    stmOutput.Close()
                End If

                Return True
            End With
        Else
            ErrMsg("File " & sFilename & " not found in index.")
        End If

        Return False

    End Function
#End If

    Private Sub ConvertPrefix(ByRef sArticle As String, ByRef sPrefix As String)
        Select Case sPrefix.ToLower
            Case ""
                sArticle = "a"
            Case "a", "an", "hers", "his", "my", "some", "the", "your"
                sArticle = sPrefix
                sPrefix = ""
            Case Else
                ' Ignore
        End Select
        If sLeft(sPrefix, 2) = "a " OrElse sLeft(sPrefix, 3) = "an " OrElse sLeft(sPrefix, 5) = "some " OrElse sLeft(sPrefix, 4) = "his " OrElse sLeft(sPrefix, 5) = "hers " OrElse sLeft(sPrefix, 4) = "the " Then
            sArticle = Split(sPrefix, " ")(0)
            sPrefix = sRight(sPrefix, Len(sPrefix) - sInstr(sPrefix, " "))
        End If
    End Sub


    Private Function GetRoomGroupFromList(ByRef iPos As Integer, ByVal salLocations As StringArrayList, ByVal sGeneratedDescription As String) As clsGroup

        Dim grp As clsGroup = Nothing
        ' Check to see if a room exists with same rooms
        For Each agrp As clsGroup In Adventure.htblGroups.Values
            If agrp.GroupType = clsGroup.GroupTypeEnum.Locations Then
                If agrp.arlMembers.Count = salLocations.Count Then
                    For Each sGroupKey As String In agrp.arlMembers
                        If Not salLocations.Contains(sGroupKey) Then GoTo NextGroup
                    Next
                    grp = agrp
                    Exit For
                End If
            End If
NextGroup:
        Next
        If grp Is Nothing Then
            grp = New clsGroup
            grp.Name = "Generated group for " & sGeneratedDescription
            For Each sLocKey As String In salLocations
                grp.arlMembers.Add(sLocKey)
            Next
            grp.Key = "GeneratedLocationGroup" & Adventure.htblGroups.Count + 1
            grp.GroupType = clsGroup.GroupTypeEnum.Locations
            Adventure.htblGroups.Add(grp, grp.Key)
        Else
            grp.Name &= " and " & sGeneratedDescription
        End If

        Return grp

    End Function


    Public Function CurrentMaxPriority(Optional ByVal bIncludeLibrary As Boolean = False) As Integer
        Dim iMax As Integer = 0

        For Each tas As clsTask In Adventure.htblTasks.Values
            If bIncludeLibrary Then
                If tas.Priority > iMax Then iMax = tas.Priority
            Else
                If tas.Priority > iMax AndAlso tas.Priority < 50000 Then iMax = tas.Priority
            End If
        Next

        Return iMax
    End Function


    Friend Sub LoadLibraries(ByVal eLoadWhat As LoadWhatEnum, Optional ByVal sOnlyLoad As String = "")
        Dim sLibraries() As String = GetSetting("ADRIFT", "Generator", "Libraries").Split("|"c)
        Dim sError As String = ""

        If sLibraries.Length = 0 OrElse (sLibraries.Length = 1 AndAlso sLibraries(0) = "") Then
            If File.Exists(BinPath & Path.DirectorySeparatorChar & "StandardLibrary.amf") Then
                ReDim sLibraries(0)
                sLibraries(0) = BinPath & Path.DirectorySeparatorChar & "StandardLibrary.amf"
            End If
        End If

        For Each sLibrary As String In sLibraries
            Dim bLoad As Boolean = True
            If sLibrary.Contains("#") Then
                bLoad = CBool(sLibrary.Split("#"c)(1))
                sLibrary = sLibrary.Split("#"c)(0)
            End If
            If bLoad AndAlso File.Exists(sLibrary) AndAlso (sOnlyLoad = "" OrElse IO.Path.GetFileNameWithoutExtension(sLibrary).ToLower = sOnlyLoad) Then
                Dim bLoadLibrary As Boolean = True
#If DEBUG Then
                bLoadLibrary = True
#End If
                If bLoadLibrary Then LoadFile(sLibrary, FileTypeEnum.XMLModule_AMF, eLoadWhat, True)
            End If
        Next

        If sError <> "" Then
            ErrMsg("Sorry.  The unregistered version of ADRIFT will only load the original library files.  The following libraries were not loaded:" & vbCrLf & vbCrLf & sError)
        End If

    End Sub


    Friend Sub OverwriteLibraries(ByVal eLoadWhat As LoadWhatEnum)

        Dim sLibraries() As String = GetSetting("ADRIFT", "Generator", "Libraries").Split("|"c)

        For Each sLibrary As String In sLibraries
            Dim bLoad As Boolean = True
            If sLibrary.Contains("#") Then
                bLoad = CBool(sLibrary.Split("#"c)(1))
                sLibrary = sLibrary.Split("#"c)(0)
            End If
            If bLoad AndAlso File.Exists(sLibrary) Then
                LoadFile(sLibrary, FileTypeEnum.XMLModule_AMF, eLoadWhat, True, Adventure.LastUpdated)
            End If
        Next

    End Sub


    ' v4 GetLine
    Private Function GetLine(ByVal bData() As Byte, ByRef iStartPos As Integer) As String

        Try
            Dim iEnd As Integer = Array.IndexOf(bData, CByte(13), iStartPos)
            If iEnd < 0 Then
                iEnd = bData.Length - 1
                If iEnd < iStartPos Then iEnd = iStartPos
            End If

            Try
                GetLine = System.Text.Encoding.Default.GetString(bData, iStartPos, iEnd - iStartPos)
            Catch
                '    GetLine = "0"
                MsgBox("iStartPos: " & iStartPos & ", iEnd: " & iEnd & ", bData.Length: " & bData.Length)
                GetLine = ""
            End Try

            iStartPos = iEnd + 2
        Catch ex As Exception
            Return ""
        End Try

    End Function


    ' Simple encryption
    Function Dencode(ByVal sText As String) As String

        Rnd(-1)
        Randomize(1976)

        Dencode = ""
        For n As Integer = 1 To sText.Length
            Dencode = Dencode & Chr((Asc(Mid(sText, n, 1)) Xor Int(CInt(Rnd() * 255 - 0.5))) Mod 256)
        Next n

    End Function

    Friend Function Dencode(ByVal sText As String, ByVal lOffset As Long) As Byte() ' String
        Rnd(-1)
        Randomize(1976)

        For n As Long = 1 To lOffset - 1
            Rnd()
        Next

        Dim result(sText.Length - 1) As Byte
        For n As Integer = 1 To sText.Length
            result(n - 1) = CByte((Asc(Mid(sText, n, 1)) Xor Int(CInt(Rnd() * 255 - 0.5))) Mod 256)
        Next

        Return result
    End Function

    Friend Function Dencode(ByVal bData As Byte(), ByVal lOffset As Long) As Byte() ' String
        Rnd(-1)
        Randomize(1976)

        For n As Long = 1 To lOffset - 1
            Rnd()
        Next

        Dim result(bData.Length - 1) As Byte
        For n As Integer = 1 To bData.Length
            'result(n - 1) = CByte((Asc(CChar(sMid(sText, n, 1))) Xor Int(CInt(Rnd() * 255 - 0.5))) Mod 256)
            result(n - 1) = CByte((bData(n - 1) Xor Int(CInt(Rnd() * 255 - 0.5))) Mod 256)
        Next

        Return result
    End Function


    Private Enum ComboEnum
        Dynamic
        WithState
        WithStateOrOpenable
        Surface
        Container
        Wearable
        Sittable
        Standable
        Lieable
    End Enum
    Private Function GetObKey(ByVal iComboIndex As Integer, ByVal eCombo As ComboEnum) As String
        Dim iMatching As Integer
        Dim i As Integer = 1
        Dim sKey As String
        Dim ob As clsObject = Nothing

        Try

            While iMatching <= iComboIndex AndAlso i < Adventure.htblObjects.Count + 1
                sKey = "Object" & i
                ob = Adventure.htblObjects(sKey)
                Select Case eCombo
                    Case ComboEnum.Dynamic
                        If Not ob.IsStatic Then iMatching += 1
                    Case ComboEnum.WithState
                        If salWithStates.Contains(sKey) Then iMatching += 1
                    Case ComboEnum.WithStateOrOpenable
                        If salWithStates.Contains(sKey) OrElse ob.Openable Then iMatching += 1
                    Case ComboEnum.Surface
                        If ob.HasSurface Then iMatching += 1
                    Case ComboEnum.Container
                        If ob.IsContainer Then iMatching += 1
                    Case ComboEnum.Wearable
                        If ob.IsWearable Then iMatching += 1
                    Case ComboEnum.Sittable
                        If ob.IsSittable Then iMatching += 1
                    Case ComboEnum.Standable
                        If ob.IsStandable Then iMatching += 1
                    Case ComboEnum.Lieable
                        If ob.IsLieable Then iMatching += 1
                End Select
                i += 1
            End While
            If ob IsNot Nothing Then
                Return ob.Key
            Else
                Return ""
            End If
        Catch ex As Exception
            ErrMsg("GetObKey error", ex)
        End Try

        Return Nothing
    End Function

End Module