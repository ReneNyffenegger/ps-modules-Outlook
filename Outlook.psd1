@{
   RootModule        = 'Outlook.psm1'
   ModuleVersion     = '0.2'
   RequiredModules   = @(
      'MS-Office'
   )
   FunctionsToExport = @(
     'send-outlookMail'
    )
}
