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

  # 2021-11-29 / V.4: Loop multiple times
    $insLoopAgain  = $true
    while ($insLoopAgain) {
       write-host 'ins loop'

       $insLoopAgain = $false
       $ins_          = $ol.inspectors

       foreach ($ins in $ins_) {
          $insLoopAgain = $true 
          write-host "   $($ins.caption)"
          $ins.close(1) # 1 = olDiscard
       }
    }


#   2021-11-29 / V.4: Loop multiple times
    $rmdLoopAgain  = $true
    while ($rmdLoopAgain) {
       write-host 'rmd loop'

       $rmdLoopAgain = $false
       $rmd_          = $ol.reminders

       foreach ($rmd in $rmd_) {
          if ($rmd.isVisible) {
             $rmdLoopAgain = $true
             write-host "Reminder: $($rmd.caption)"
             $rmd.dismiss()
          }
          else {
          #  write-host "$($rmd.caption) is not visible"
          }
       }
    }
}

function disable-outlookNotifications {
   $regKeyOfficeRootV = get-msOfficeRegRoot

   $regKeyOfficeRootV

   set-itemProperty "$regKeyOfficeRootV\outlook\preferences" -name newMailDesktopAlerts -type dWord -value 0

}
