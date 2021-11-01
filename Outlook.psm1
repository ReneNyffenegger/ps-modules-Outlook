set-strictMode -version latest

function send-outlookMail {

   param(
      [parameter(mandatory = $true)]
      [string]    $recipient
   ,
      [parameter(mandatory = $true)]
      [string]    $subject
   ,
      [parameter(mandatory = $true)]
      [string]    $body
   ,
      [parameter(mandatory = $false)]
      [string[]]  $attachments
   )

#   $ol   = get-activeObject outlook.application
    $ol   = get-msOfficeComObject  outlook

    $email = $ol.createItem(0)  # 0 = olMailItem
    $email.to      = $recipient
    $email.subject = $subject
    $email.body    = $body

    foreach ($attachment in $attachments) {
    $resolved_path = resolve-path $attachment
       if (! (test-path $resolved_path)) {
          write-host "Attachment $resolved_path was not found"
          return
       }
     #
     # For some reason, apparently, resolved path
     # must be put into double quotes because otherwise,
     # the error
     #     Value does not fall within the expected range.
     # is thrown (which, imho, does not make lot of sense).
     #
       $null = $email.attachments.add("$resolved_path")
    }
 #  $email.display()
    $email.send()
}

function close-outlookWindows {
    $ol = get-msOfficeComObject outlook

    foreach ($ins in $ol.inspectors) {
       $ins.close(1) # 1 = olDiscard
    }

    foreach ($rmd in $ol.reminders) {
       if ($rmd.isVisible) {
          write-host "$($rmd.caption) is visible"
          $rmd.dismiss()
       }
       else {
          write-host "$($rmd.caption) is not visible"
       }
    }
}
