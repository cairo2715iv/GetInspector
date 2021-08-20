# GetInspector
  $oItem.GetInspector    $sBody =  $chem_modele_html   $oItem.HTMLBody = $sBody
 ; Add all attachments
   If $pj1<>"" Then
      $oItem = _OL_ItemAttachmentAdd($oOutlook, $oItem, Default, $pj1)
      If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", $pj1 & "Error adding an attachment to a mail in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = "  & @error & ", @extended = " & @extended)
   EndIf

   If $pj2<>"" Then
      $oItem = _OL_ItemAttachmentAdd($oOutlook, $oItem, Default, $pj2)
      If @error <> 0 Then Exit MsgBox(16, "OutlookEX UDF: _OL_ItemCreate Example Script", $pj2 & "Error adding an attachment to a mail in folder 'Outlook-UDF-Test\TargetFolder\Mail'. @error = "  & @error & ", @extended = " & @extended)
   EndIf
