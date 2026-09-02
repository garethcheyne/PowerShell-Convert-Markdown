@{
    RootModule           = 'ConvertMarkdown.psm1'

    # Versioned yyyy.MM.dd. [version] normalises away zero padding
    # (2026.08.29 -> 2026.8.29); the padded string is in PrivateData.DisplayVersion
    # and is what Get-ConvertMarkdownVersion reports.
    ModuleVersion        = '2026.9.2'

    GUID                 = 'c5e81f47-9a2d-4b63-8f10-6d3e7c94a2b8'
    Author               = 'Gareth Cheyne'
    CompanyName          = 'Gareth Cheyne'
    Copyright            = '(c) 2026 Gareth Cheyne. MIT Licence.'
    Description          = 'Converts Markdown into styled Word .docx and PDF documents, with mermaid diagram rendering.'

    PowerShellVersion    = '7.0'
    CompatiblePSEditions = @('Core')

    FunctionsToExport    = @('Convert-Markdown', 'New-ConvertMarkdownStyle', 'Update-ConvertMarkdown', 'Get-ConvertMarkdownVersion')
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @('md2docx', 'md2pdf', 'Convert-MarkdownToWord')

    PrivateData          = @{
        DisplayVersion = '2026.09.02'
        LicenseUri     = 'https://github.com/garethcheyne/PowerShell-Convert-Markdown/blob/main/LICENSE'
        ProjectUri     = 'https://github.com/garethcheyne/PowerShell-Convert-Markdown'

        PSData         = @{
            Tags         = @('Markdown', 'Word', 'docx', 'PDF', 'Mermaid', 'Documentation')
            ReleaseNotes = 'Consolidates seven drifted copies of Convert-MarkdownToWord.ps1. Adds PDF output, 10mm margins, installable module, yyyy.MM.dd versioning.'
        }
    }
}
