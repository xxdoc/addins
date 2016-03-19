Attribute VB_Name = "mReservedWords"
Option Explicit

Public KeyWords             As Collection

'new keywords
Public Const sApiConst      As String = "ApiConst "
Public Const sApiDeclare    As String = "ApiDeclare "
Public Const sApiType       As String = "ApiType "

'the ReservedWords come from VB6.exe, they are in there (most of them anyway)
Public Const ResWds As String = ".#Const.#Else.#ElseIf.#End.#If.Abs.Access.AddressOf.Alias.And.Any.Append.Array.As.Assert." & _
                                "Base.Binary.Boolean.ByRef.ByVal.Call.Case.CBool.CByte.CCur.CDate.CDbl.CDec.CDecl." & _
                                "ChDir.ChDrive.CInt.Choose.Circle.CLng.Close.Compare.Const.CSng.CStr.CurDir.Currency.CVar.CVDate." & _
                                "CVErr.Database.Date.Date.Debug.Declare.DefBool.DefByte.DefCur.DefDate.DefDbl.DefDec." & _
                                "DefInt.DefLng.DefObj.DefSng.DefStr.DefVar.Dim.Dir.Do.DoEvents.Double.Each.Else.ElseIf." & _
                                "End.Enum.Eqv.Erase.Err.Error.Event.Exit.Explicit.False.Fix.For.Format.FreeFile.Friend.Function.Get." & _
                                "Global.GoSub.GoTo.Hide.If.IIf.Imp.Implements.In.Input.InputB.InStr.InStrB.InstrRev.Int.Integer.Is.LBound." & _
                                "Left.Len.LenB.Let.Lib.Like.Line.Load.Local.Lock.Long.Loop.LSet.Mid.MidB.Mod.Module.Name." & _
                                "New.Next.Not.Nothing.Object.ObjPtr.On.Open.Option.Optional.Or.Output.ParamArray.Preserve." & _
                                "Print.Private.Property.PSet.Public.Put.RaiseEvent.Random.Randomize.Read.ReDim.Rem.'." & _
                                "Replace.Reset.Resume.Return.RGB.Right.Rnd.RSet.Scale.Seek.Select.Set.Sgn.Shared.Show.Single.Spc." & _
                                "Static.Step.Stop.StrComp.String.StrPtr.Sub.Switch.Tab.Then.To.True.Type.TypeOf.UBound.Unknown." & _
                                "Unload.Unlock.Until.Variant.VarPtr.VarType.Wend.While.Width.With.WithEvents.Write.Xor."

Public Const OthWds As String = ".AppActivate.Clipboard.Command$.CommitTrans.CompactDatabase.CreateObject." & _
                                "DateAdd.DateDiff.DatePart.DateSerial.DateValue.DeleteSetting.Document.Environ$." & _
                                "FileAttr.FileCopy.FileDateTime.FileLen.Format$.FreeLocks.GetAllSettings.GetAttr." & _
                                "GetObject.GetSetting.Hour.InputBox.IsArray.IsDate.IsEmpty.IsError.IsMissing.IsNull." & _
                                "IsNumeric.IsObject.LCase$.Left$.LoadPicture.LoadResData.LoadResPicture.LoadResString." & _
                                "LTrim$.Mid$.Minute.Month.MsgBox.Partition.QBColor.Replace$.Right$.Rollback.RTrim$.SavePicture." & _
                                "SaveSetting.Screen.Second.SendKeys.SetAttr.SetDataAccessOption.SetDefaultWorkpace." & _
                                "Space$.String$.Switch.Timer.TimeValue.Trim$.TypeName.UBound.UCase$.vbArrowHourglass." & _
                                "vbBlack.vbBlue.vbChecked.vbCrLf.vbCyan.vbDefault.vbFormFeed.vbGrayed.vbGreen.vbHourglass." & _
                                "vbMagenta.vbMaximized.vbMinimized.vbModal.vbNormal.vbNullString.vbObjectError.vbPixels." & _
                                "vbRed.vbSrcCopy.vbTwips.vbUnchecked.vbWhite.vbYellow.Weekday.Year."

Public Sub LoadKeywords(LongerThan As Long)

  'create the keywords collection

  Dim Words()   As String
  Dim Index     As Long

    Set KeyWords = Nothing
    Set KeyWords = New Collection

    With KeyWords
        If Not ApiDatabase Is Nothing Then 'api database is available
            .Add RTrim$(sApiConst)
            .Add RTrim$(sApiDeclare)
            .Add RTrim$(sApiType)
        End If

        Words = Split(OthWds, ".")
        For Index = 0 To UBound(Words)
            If Len(Words(Index)) > LongerThan Then
                .Add Words(Index)
            End If
        Next Index
        Words = Split(ResWds, ".")
        For Index = 0 To UBound(Words)
            If Len(Words(Index)) > LongerThan Then
                .Add Words(Index)
            End If
        Next Index
    End With 'KEYWORDS

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 16:14)  Decl: 35  Code: 36  Total: 71 Lines
':) CommentOnly: 3 (4,2%)  Commented: 2 (2,8%)  Empty: 11 (15,5%)  Max Logic Depth: 4
