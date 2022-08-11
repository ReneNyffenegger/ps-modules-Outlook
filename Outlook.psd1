@{
   RootModule        = 'Outlook.psm1'
   ModuleVersion     = '0.5'
   RequiredModules   = @(
      'MS-Office'
   )
   FunctionsToExport = @(
     'send-outlookMail',
     'close-outlookWindows'
     'disable-outlookNotifications'
    )
}
