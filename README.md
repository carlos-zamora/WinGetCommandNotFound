# WinGetCommandNotFound

This is a proof-of-concept implementing both the `IFeedbackProvider` interfaces.
`IFeedbackProvider` requires PS7.4+.

## Feedback provider

The feedback provider uses WinGet's `Microsoft.Management.Deployment.winmd` as well as their `winrtact.dll`. These can be found in the `Microsoft.WinGet.Client-PSModule.zip` asset from the [WinGet CLI releases page](https://github.com/microsoft/winget-cli/releases).

## Building

Go to `src` folder and use `dotnet build`.  Requires .NET 8 SDK installed and in path.

## Using

Copy `winrtact.dll` from `Microsoft.WinGet.Client-PSModule.zip` to the published folder after building. Then, just `Import-Module WinGetCommandNotFound.psd1` which will register the Feedback Provider.
