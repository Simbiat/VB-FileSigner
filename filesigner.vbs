function filesigner(moveonly, signonly)
  if silent = true Then
      On Error Resume Next
  end if
  Dim paramlist, parami, listitem, fakeloop, filesfound, curdate, signtemp, signreq, actiontype
  Dim source_path, filemask, out_path, bck_source, bck_result, sign_flag, encrypt_flag, unsign_flag, decrypt_flag, sign_profile, crypt_subject
  Dim repfolder, repmask, report, repname, signaturapath, sigcom, fulllog, objTextFile
  Dim backorig, backresu, backorigi, backresui, expanddir, tmpdir, signerr, errordisc
  'checking if both paramters are set to true
  if moveonly = true and signonly = true then
      'exiting, since this means that we do not need to do anything
      exit function
  end if
  repname = ""
  signtemp = ReadIni("SIGNER", "signtemp")
  curdate = dateyyymmdd()
  Set repmask = New RegExp
  repmask.IgnoreCase = True
  paramlist = split(ReadIni("SIGNER", "main"), ";;;")
  filesfound = false
  signreq = false
  statusbar.value = "Searching for files..."
  for parami=Lbound(paramlist) to ubound(paramlist)
      listitem = split(paramlist(parami), ":::")
      source_path = listitem(0)
      sign_flag = listitem(5)
      encrypt_flag = listitem(6)
      unsign_flag = listitem(7)
      decrypt_flag = listitem(8)
      sign_profile = listitem(9)
      if trim(source_path) = "" Then
          exit for
      End if
      if ReadIni("SIGNER", trim(source_path)) <> "" Then
          source_path = ReadIni("SIGNER", trim(source_path))
      end if
      if Right(source_path, 1) <> "\" then
          exit for
      end if
      source_path = Replace(source_path, "YYYYMMDD", curdate)
      If not objfso.FolderExists(source_path) Then
          exit for
      else
          repmask.Pattern = listitem(1)
          set repfolder=objFSO.GetFolder(source_path)
          for each report in repfolder.Files
              repname = objFSO.GetFileName(report)
              if repmask.Test(repname) Then
                  filesfound = true
                  if (sign_flag = 1 or encrypt_flag = 1 or unsign_flag = 1 or decrypt_flag = 1) and trim(sign_profile) <> "" then
                      signreq = true
                  end if
                  Exit For
              end if
          next
          if signreq = true then
              Exit for
          end if
      end if
  next
  if filesfound = false then
      call logwrite (HTA_Log, "[Info   ]" & logstartline() & "No files found", "No files found", ForAppending, 1, 0)
      Exit Function
  end if
  if moveonly = true then
      signreq = false
  end if
  if signreq = true and moveonly = false then
      signaturapath = signaturacheck()
      If signaturapath = "" Then
          call logwrite (HTA_Log, "[SignErr]" & logstartline() & "No Signatura client found!", "No Signatura client found!", ForAppending, 3, 0)
          Exit Function
      end if
      if hashfilecheck(ReadIni("Files", "hashdir") & objNet.ComputerName & "_signatura.txt") = false Then
          Exit Function
      end if
      if not isElevated Then
          call logwrite (HTA_Log, "[SignErr]" & logstartline() & "Process started with no elevation!", "Process started with no elevation!" & vbcrlf & "Only files not requiring Sgnatura will be processed!", ForAppending, 3, 0)
          signreq = false
      end if
      call filemover(ReadIni("SIGNER", "pkiconf"), Replace(signaturapath, "spki1utl.exe", "pki1.conf"), 0, 1)
  end if
  for parami=Lbound(paramlist) to ubound(paramlist)
      for fakeloop=1 to 1
          listitem = split(paramlist(parami), ":::")
          source_path = listitem(0)
          filemask = listitem(1)
          out_path = listitem(2)
          bck_source = listitem(3)
          bck_result = listitem(4)
          sign_flag = listitem(5)
          encrypt_flag = listitem(6)
          unsign_flag = listitem(7)
          decrypt_flag = listitem(8)
          sign_profile = listitem(9)
          crypt_subject = listitem(10)
          'Check for need to sign on earlier for performance and if out_path is " " (space)
          'if (sign_flag = 1 or encrypt_flag = 1 or unsign_flag = 1 or decrypt_flag = 1) and (signreq = false or out_path = " ") then
          if (sign_flag = 1 or encrypt_flag = 1 or unsign_flag = 1 or decrypt_flag = 1) and (signreq = false) then
              exit for
          end if
          'check if we are here just to sign
          if sign_flag = 0 and encrypt_flag = 0 and unsign_flag = 0 and decrypt_flag = 0 and signonly = true then
              exit for
          end if
          'folders validation and expansion
          if trim(source_path) = "" Then
              exit for
          End if
          'if trim(out_path) = "" Then
          '    exit for
          'End if
          if ReadIni("SIGNER", trim(source_path)) <> "" Then
              source_path = ReadIni("SIGNER", trim(source_path))
          end if
          if out_path <> " " and ReadIni("SIGNER", trim(out_path)) <> "" Then
              out_path = ReadIni("SIGNER", trim(out_path))
          end if
          if Right(source_path, 1) <> "\" then
              exit for
          end if
          if out_path <> " " and Right(out_path, 1) <> "\" then
              exit for
          end if
          source_path = Replace(source_path, "YYYYMMDD", curdate)
          if out_path <> " " then
              out_path = Replace(out_path, "out_path", curdate)
              If not objfso.FolderExists(out_path) Then
                  On Error Resume Next
                  objFSO.CreateFolder(out_path)
                  On Error Goto 0
                  If not objfso.FolderExists(out_path) Then
                      exit for
                  end if
              end if
          end if
          If not objfso.FolderExists(source_path) Then
              exit for
          end if
          if trim(bck_source) <> "" Then
              backorig = split(bck_source, "|")
              expanddir = false
              for backorigi=Lbound(backorig) to ubound(backorig)
                  tmpdir = backorig(backorigi)
                  if ReadIni("SIGNER", trim(tmpdir)) <> "" Then
                      tmpdir = ReadIni("SIGNER", trim(tmpdir))
                  end if
                  if Right(tmpdir, 1) <> "\" then
                      expanddir = true
                      exit for
                  end if
                  tmpdir = Replace(tmpdir, "YYYYMMDD", curdate)
                  If not objfso.FolderExists(tmpdir) Then
                      On Error Resume Next
                      objFSO.CreateFolder(tmpdir)
                      On Error Goto 0
                      If not objfso.FolderExists(tmpdir) Then
                          expanddir = true
                          exit for
                      end if
                  end if
              Next
              if expanddir = true then
                  exit for
              end if
          end if
          if trim(bck_result) <> "" Then
              backresu = split(bck_result, "|")
              expanddir = false
              for backresui=Lbound(backresu) to ubound(backresu)
                  tmpdir = backresu(backresui)
                  if ReadIni("SIGNER", trim(tmpdir)) <> "" Then
                      tmpdir = ReadIni("SIGNER", trim(tmpdir))
                  end if
                  if Right(tmpdir, 1) <> "\" then
                      expanddir = true
                      exit for
                  end if
                  tmpdir = Replace(tmpdir, "YYYYMMDD", curdate)
                  If not objfso.FolderExists(tmpdir) Then
                      On Error Resume Next
                      objFSO.CreateFolder(tmpdir)
                      On Error Goto 0
                      If not objfso.FolderExists(tmpdir) Then
                          expanddir = true
                          exit for
                      end if
                  end if
              Next
              if expanddir = true then
                  exit for
              end if
          end if
          'validations
          if trim(filemask) = "" Then
              exit for
          End if
          if (sign_flag <> 0 and sign_flag <> 1) or (encrypt_flag <> 0 and encrypt_flag <> 1) or (unsign_flag <> 0 and unsign_flag <> 1) or (decrypt_flag <> 0 and decrypt_flag <> 1) then
              exit for
          end if
          if trim(crypt_subject) <> "" then
              crypt_subject = ReadIni("SIGNER", crypt_subject)
          end if
          if encrypt_flag = 1 and crypt_subject = "" then
              exit for
          end if
          if (sign_flag = 1 or encrypt_flag = 1 or unsign_flag = 1 or decrypt_flag = 1) and trim(sign_profile) = "" then
              exit for
          end if
          if sign_flag = 1 or encrypt_flag = 1 then
              unsign_flag = 0
              decrypt_flag = 0
          end if
          repmask.Pattern = filemask
          set repfolder=objFSO.GetFolder(source_path)
          for each report in repfolder.Files
              repname = objFSO.GetFileName(report)
              statusbar.value = "Processing " & repname & "..."
              if repmask.Test(repname) Then
                  sigcom = ""
                  if sign_flag = 1 and encrypt_flag = 1 then
                      actiontype = "signed and encrypted"
                      sigcom = "cmd /c ""cd /D """& Replace(signaturapath, "spki1utl.exe", "") &"""&&""" & signaturapath & """ -profile """ & sign_profile & """ -sign -encrypt -in """ & signtemp & repname & """ -out """ & signtemp & repname & "_signed"" -recsubj """ & crypt_subject & """ > """ & signtemp & repname & ".log" & """ 2>&1"""
                  elseif sign_flag = 1 and encrypt_flag = 0 then
                      actiontype = "signed"
                      sigcom = "cmd /c ""cd /D """& Replace(signaturapath, "spki1utl.exe", "") &"""&&""" & signaturapath & """ -profile """ & sign_profile & """ -sign -data """ & signtemp & repname & """ -out """ & signtemp & repname & "_signed"" > """ & signtemp & repname & ".log" & """ 2>&1"""
                  elseif sign_flag = 0 and encrypt_flag = 1 then
                      actiontype = "encrypted"
                      sigcom = "cmd /c ""cd /D """& Replace(signaturapath, "spki1utl.exe", "") &"""&&""" & signaturapath & """ -profile """ & sign_profile & """ -encrypt -in """ & signtemp & repname & """ -out """ & signtemp & repname & "_signed"" -recsubj """ & crypt_subject & """ > """ & signtemp & repname & ".log" & """ 2>&1"""
                  else
                      if unsign_flag = 1 and decrypt_flag = 1 then
                          actiontype = "unsigned and decrypted"
                          sigcom = "cmd /c ""cd /D """& Replace(signaturapath, "spki1utl.exe", "") &"""&&""" & signaturapath & """ -profile """ & sign_profile & """ -decrypt -verify -in """ & signtemp & repname & """ -out """ & signtemp & repname & "_signed"" -delete -1> """ & signtemp & repname & ".log" & """ 2>&1"""
                      elseif unsign_flag = 1 and decrypt_flag = 0 then
                          actiontype = "unsigned"
                          sigcom = "cmd /c ""cd /D """& Replace(signaturapath, "spki1utl.exe", "") &"""&&""" & signaturapath & """ -profile """ & sign_profile & """ -verify -in """ & signtemp & repname & """ -out """ & signtemp & repname & "_signed"" -delete -1> """ & signtemp & repname & ".log" & """ 2>&1"""
                      elseif unsign_flag = 0 and decrypt_flag = 1 then
                          actiontype = "decrypted"
                          sigcom = "cmd /c ""cd /D """& Replace(signaturapath, "spki1utl.exe", "") &"""&&""" & signaturapath & """ -profile """ & sign_profile & """ -decrypt -in """ & signtemp & repname & """ -out """ & signtemp & repname & "_signed""> """ & signtemp & repname & ".log" & """ 2>&1"""
                      end if
                  end if
                  if sigcom = "" then
                      if trim(bck_source) <> "" Then
                          backorig = split(bck_source, "|")
                          for backorigi=Lbound(backorig) to ubound(backorig)
                              tmpdir = backorig(backorigi)
                              if ReadIni("SIGNER", trim(tmpdir)) <> "" Then
                                  tmpdir = ReadIni("SIGNER", trim(tmpdir))
                              end if
                              tmpdir = Replace(tmpdir, "YYYYMMDD", curdate)
                              call filemover(source_path & repname, tmpdir & repname, 0, 0)
                          Next
                      end if
                      if out_path <> " " then
                          call filemover(source_path & repname, out_path & repname, 1, 1)
                      else
                          call filedel(source_path & repname, 1, 0)
                      end if
                  else
                      call filemover(source_path & repname, signtemp & repname, 0, 1)
                      statusbar.value = "Processing " & repname & "..."
                      signerr = oShell.Run (sigcom, 1, True)
                      If objFSO.fileexists(signtemp & repname & ".log") Then
                          Set objTextFile = objFSO.openTextfile(signtemp & repname & ".log", ForReading)
                          fulllog = StrConv(objTextFile.readall, "Windows-1251", "cp866")
                          if cstr(hex(signerr)) <> "0" Then
                              errordisc = Trim(Replace(Mid(fulllog, Instr(fulllog, "ђҐ§г«мв в: ") + 11, len(fulllog) - Instr(fulllog, "ђҐ§г«мв в: ") - 22), vbCrLf, ""))
                          end if
                          fulllog = Replace(fulllog, vbCrLf, "<br>")
                          if cstr(hex(signerr)) <> "0" Then
                              call logwrite (HTA_Log, "[SignErr]" & logstartline() & fulllog, 0, ForAppending, 0, 0)
                          else
                              call logwrite (HTA_Log, "[SignSuc]" & logstartline() & fulllog, 0, ForAppending, 0, 0)
                          end if
                          objTextFile.close
                          call filedel(signtemp & repname & ".log", 1, 1)
                      end if
                      if not objFSO.fileexists(signtemp & repname & "_signed") or cstr(hex(signerr)) <> "0" Then
                          if objFSO.fileexists(signtemp & repname & "_signed") Then
                              call filedel(signtemp & repname & "_signed", 1, 1)
                          end if
                          call filedel(signtemp & repname, 1, 1)
                          call logwrite (HTA_Log, "[SignErr]" & logstartline() & signtemp & repname & " failed to get " & actiontype & " with Signatura", signtemp & repname & " failed to get " & actiontype & " with Signatura with error" & vbCrLf & """" & errordisc & """", ForAppending, 3, 0)
                      else
                          call filedel(signtemp & repname, 1, 1)
                          if trim(bck_source) <> "" Then
                              backorig = split(bck_source, "|")
                              for backorigi=Lbound(backorig) to ubound(backorig)
                                  tmpdir = backorig(backorigi)
                                  if ReadIni("SIGNER", trim(tmpdir)) <> "" Then
                                      tmpdir = ReadIni("SIGNER", trim(tmpdir))
                                  end if
                                  tmpdir = Replace(tmpdir, "YYYYMMDD", curdate)
                                  call filemover(source_path & repname, tmpdir & repname, 0, 0)
                              Next
                          end if
                          if trim(bck_result) <> "" Then
                              backresu = split(bck_result, "|")
                              for backresui=Lbound(backresu) to ubound(backresu)
                                  tmpdir = backresu(backresui)
                                  if ReadIni("SIGNER", trim(tmpdir)) <> "" Then
                                      tmpdir = ReadIni("SIGNER", trim(tmpdir))
                                  end if
                                  tmpdir = Replace(tmpdir, "YYYYMMDD", curdate)
                                  call filemover(signtemp & repname & "_signed", tmpdir & repname, 0, 0)
                              Next
                          end if
                          call filemover(signtemp & repname & "_signed", out_path & repname, 1, 1)
                          call filedel(source_path & repname, 1, 0)
                          call logwrite (HTA_Log, "[SignSuc]" & logstartline() & repname & " " & actiontype & " by Signatura", repname & " " & actiontype & " by Signatura", ForAppending, 1, 0)
                      end if
                  end if
              end if
          next
      Next
  Next
  if signreq = true and moveonly = false then
      call filedel(Replace(signaturapath, "spki1utl.exe", "pki1.conf"), 1, 1)
  end if
  statusbar.value = "Searching for files completed"
  if silent = true Then
      On Error Goto 0
  end if
end function
