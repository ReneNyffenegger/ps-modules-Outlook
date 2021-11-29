@{
   RootModule        = 'Outlook.psm1'
   ModuleVersion     = '0.4'
   RequiredModules   = @(
      'MS-Office'
   )
   FunctionsToExport = @(
     'send-outlookMail',
     'close-outlookWindows'
    )
}
