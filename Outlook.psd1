@{
   RootModule        = 'Outlook.psm1'
   ModuleVersion     = '0.3'
   RequiredModules   = @(
      'MS-Office'
   )
   FunctionsToExport = @(
     'send-outlookMail',
     'close-outlookWindows'
    )
}
