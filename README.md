# CmdletJack
CmdletJack creates spreadsheets (.CSV files) of cmdlet documentation
data as reports. Usage is the path to the directory of the cmdlet
projects or to a specific file. Reports are automatically generated.
Syntax for all projects:

    CmdletJack <C:\Path\Cmdlets>

Syntax for a specifc project:

    CmdletJack <C:\Path\Cmdlets\VMM.XML>

Existing reports are overwritten, unless you follow with '-a' to append.

You can also search:

    CmdletJack <C:\Path\Cmdlets> search <pattern> [<matchcase|wholeword|both>]

Search options: specify 'matchcase' to match case, 'wholeword' to match whole words,
or 'both' to use matchcase and wholeword.

CmdletJack version: 2.7 - (c) 2016 Microsoft Corporation. All rights reserved.
