<#
.SYNOPSIS
    Title   : Show-Output-Exception-ErrorTrace-Info.ps1
    Author  : Jon Damvi
    Version : 1.0.0
    Date    : 19.05.2025
    License : MIT

   Release Notes: v1.0.0 (19.05.2025) - initial release.

.DESCRIPTION
    Shows detailed Exception Error Trace information

.INPUTS
    None.

.OUTPUTS
    None.

#>

class ActualError {
    [int]$HResult
    [string]$HResultHex
    [string]$Facility
    [int]$FacilityCode
    [int]$ErrorCode
    [bool]$IsFailure
    [string]$Message
}

function Get-ActualError {
    Param([int]$hresult)
    # Define a class to hold the decoded error info
    class ActualError {
        [int]$HResult
        [string]$HResultHex
        [string]$Facility
        [int]$FacilityCode
        [int]$ErrorCode
        [bool]$IsFailure
        [string]$Message
    }
    $actualError = [ActualError]::new()
    $actualError.HResult = $hresult
    $actualError.HResultHex = "0x" + $hresult.ToString("X8")
    # Extract error code (lowest 16 bits)
    $actualError.ErrorCode = $hresult -band 0xFFFF
    # Extract facility code (bits 16-30)
    $actualError.FacilityCode = ($hresult -shr 16) -band 0x1FFF
    # Check if severity bit (bit 31) is set (failure)
    $actualError.IsFailure = ($hresult -band 0x80000000) -ne 0
    # Get the Win32 error message corresponding to the HRESULT
    Try {
        $actualError.Message = (New-Object System.ComponentModel.Win32Exception($hresult)).Message
    } Catch {
        $actualError.Message = "Unknown error message"
    }
    # Map facility codes to names based on standard HRESULT facility codes
    switch ($actualError.FacilityCode) {
        0   { $actualError.Facility = 'FACILITY_NULL' }
        1   { $actualError.Facility = 'FACILITY_RPC' }
        2   { $actualError.Facility = 'FACILITY_DISPATCH' }
        3   { $actualError.Facility = 'FACILITY_STORAGE' }
        4   { $actualError.Facility = 'FACILITY_ITF' }
        5   { $actualError.Facility = 'FACILITY_WIN32' }
        6   { $actualError.Facility = 'FACILITY_WINDOWS' }
        7   { $actualError.Facility = 'FACILITY_SSPI' }
        default { $actualError.Facility = 'UNKNOWN' }
    }
    return $actualError
}

Function Get-ActiveScopeCount {
    [byte]$count = 0
    while ($true) {
        Try {
            # Try to get reserved variable at scope $count. This function call limits processing overhead to minimum as possible.
            Get-Variable -Name 'PID' -ValueOnly -Scope $count -ErrorAction Stop -WarningAction Ignore -InformationAction Ignore -Verbose:$false | Out-Null
            $count++
        }
        Catch {
            break
        }
    }
    return [byte]($count-2)
}

Function Get-LocalSetVariables {
    Param(
        [Parameter(Mandatory)]
        [ValidateRange(0,255)]
        [byte]$Scope = 0
    )
    # Scope = 0 - value is for the user
    # ActualScope is one level back from 0
    $ActualScope = $Scope + 1
    If($ActualScope -ge $(Get-ActiveScopeCount)) {
        return
    }
    # Known automatic variables to exclude
    $automaticVars = @(
        'null', 'PSBoundParameters', 'PSCommandPath', 'PSCulture', 'PSDefaultParameterValues',
        'PSEmailServer', 'PSHome', 'PSItem', 'PSModuleAutoLoadingPreference', 'PSScriptRoot', 'PSSessionApplicationName',
        'PSSessionConfigurationName', 'PSSessionOption', 'PSUICulture', 'PSVersionTable', 'StackTrace', 'This',
        'VerbosePreference', 'WarningPreference', 'DebugPreference', 'ErrorActionPreference', 'InformationPreference',
        'ConfirmPreference', 'WhatIfPreference', 'MaximumAliasCount', 'MaximumDriveCount', 'MaximumErrorCount',
        'MaximumFunctionCount', 'MaximumHistoryCount', 'MaximumVariableCount', 'NestedPromptLevel', 'OutputEncoding',
        'ProgressPreference', 'PWD', 'Home', 'Host', 'PID', 'ExecutionContext', 'MyInvocation', 'PSDefaultParameterValues',
        'PSModulePath', 'PSProvider', 'PSCommandPath', 'PSCulture', 'PSHOME', 'PSUICulture', 'PSCmdlet', '_'
    )
    # Flags for Constant and AllScope
    $excludeFlags = 2 -bor 8
    # Get local variables excluding Constant and AllScope
    $localVars = @(Get-Variable -Scope $ActualScope | ? { ([int]$_.Options -band $excludeFlags) -eq 0 })
    # Exclude known automatic variables by name
    $filteredVars = $localVars.GetEnumerator() | ? { $automaticVars -notcontains $_.Name }
    $result = $filteredVars.GetEnumerator() | ? {
        $parentVar = Get-Variable -Name $_.Name -Scope Script -ErrorAction SilentlyContinue
        If ($null -eq $parentVar) {
            $true
        }
        Else {
            $_.Value -ne $parentVar.Value
        }
    }
    return $result
}

Function Format-DebugInfo {
    [CmdletBinding(DefaultParameterSetName = 'TransitivePoint')]
    Param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory, ParameterSetName = 'EntryPoint')]
        [Alias('Entry')]
        [switch]$EntryPoint,

        [Parameter(Mandatory, ParameterSetName = 'TransitivePoint')]
        [Alias('Transitive')]
        [switch]$TransitivePoint,

        [Parameter()]
        [ValidateRange(0,255)]
        [byte]$Scope = 0
    )
    Try {
        $ActualScope = $Scope + 1
        If($ActualScope -ge $(Get-ActiveScopeCount)) {
            return
        }
        # Function body as before, using $PSCmdlet.ParameterSetName to branch logic
        $callStack = @(Get-PSCallStack)
        $CallStackScope = $ActualScope
        If($callStack.Count -le $CallStackScope) {
            return
        }
        $currentFunction = $callStack[$CallStackScope].FunctionName
        $invocationName  = $callStack[$CallStackScope].InvocationInfo.InvocationName
        $aliasUsed = ($invocationName -And $invocationName -ne $currentFunction)
        $EntryPointName = "${currentFunction}"
        If ($aliasUsed) {
            $EntryPointName = "#$invocationName->[${currentFunction}]:"
        }
        # Build call stack string only if in Params set
        $callStackStr = ''
        If ($PSCmdlet.ParameterSetName -eq 'EntryPoint') {
            $filteredStack = @($callStack[$CallStackScope..($callStack.Count-1)] | ForEach-Object {
                If ($_.FunctionName -eq "<ScriptBlock>") { "<Script>" }
                Else {
                    If ($_.InvocationInfo.InvocationName -And $_.InvocationInfo.InvocationName -ne $_.FunctionName -And $_.FunctionName -notmatch '.+<(Begin|Process|End)>') {
                        "#$([string]$_.InvocationInfo.InvocationName)->[$([string]$_.FunctionName)]"
                    }
                    Else {
                        [string]$_.FunctionName
                    }
                }
            })
            [Array]::Reverse($filteredStack)
            $callStackStr = $filteredStack -join ':'
        }
        # Format items (Params or Locals)
        $Params = $callStack[$CallStackScope].InvocationInfo.BoundParameters
        switch ($PSCmdlet.ParameterSetName) {
            'EntryPoint' {
                $items = $Params
            }
            'TransitivePoint' {
                $items = @{}
                $VariableScope  = $ActualScope
                $LocalVariables = @(Get-LocalSetVariables -Scope ($VariableScope)) | ? { (-Not $Params.ContainsKey($_.Name) -OR $Params[$_.Name] -ne $_.Value) }
                Foreach ($var in $LocalVariables) {
                    $items[$var.Name] = $var.Value
                }
            }
            default {
                $items = @{}
            }
        }
        # Format items using your Format-Value helper
        $itemStrings = @(Foreach ($key in $items.Keys | Sort) {
            $Value = $items[$key]
            "${key}=" + (Format-Value $Value)
        })
        $itemsStr = If ($itemStrings.Count) { $itemStrings -join ', ' } Else { '' }
        # Determine line number if call stack available
        $lineNum = $callStack[$CallStackScope].ScriptLineNumber
        If ($PSCmdlet.ParameterSetName -eq 'EntryPoint') {
            $itemsStr = "($itemsStr)"
        }
        $debugMsg = ''
        If ($PSCmdlet.ParameterSetName -eq 'EntryPoint') {
            $debugMsg = "[${EntryPointName}]: '${Message}' (Line:${lineNum}): ${callStackStr}${itemsStr}"
        } Else {
            $debugMsg = "[${EntryPointName}]: '${Message}' (Line:${lineNum}): {${callStackStr}${itemsStr}}"
        }
        return $debugMsg
    }
    Catch {
        Write-ErrorLog $_
    }
}

Function Format-Value {
    Param([object]$Value)
    $MaxLength = 256
    If ($null -eq $Value) {
        return '`$null'
    }
    elseIf ($Value -is [switch]) {
        return $Value.ToString().ToLower()
    }
    elseIf ($Value -is [bool]) {
        return $Value.ToString().ToLower()
    }
    elseIf ($Value -is [string] -OR $Value -is [ScriptBlock]) {
        If($Value -is [scriptblock]) { $Value = $Value.ToString() }
        $str = $Value.TrimStart(" ", "`n", "`r")
        # Get up to first newline or $MaxLength, whichever is shorter
        $newlineIdx = $str.IndexOf("`n")
        If ($newlineIdx -ge 0 -And $newlineIdx -lt $MaxLength) {
            $preview = $str.Substring(0, $newlineIdx).TrimEnd(" ", "`r", "`n")
            return '"' + $preview + ' \\~..."'
        } ElseIf ($str.Length -gt $MaxLength) {
            $preview = $str.Substring(0, $MaxLength).TrimEnd(" ", "`r", "`n")
            return '"' + $preview + ' ~..."'
        } Else {
            return '"' + $str + '"'
        }
    }
    elseIf ($Value.GetType().IsPrimitive) {
        return $Value.ToString()
    }
    elseIf ($Value -is [System.Array]) {
        $typeName = $Value.GetType().GetElementType().Name
        $count = $Value.Count
        return "[$typeName[$count]]`$Obj"
    }
    elseIf ($Value.ToString() -ne $Value.GetType().FullName) {
        return "'" + $Value.ToString() + "'"
    }
    Else {
        $typeName = $Value.GetType().Name
        return "[$typeName]`$Obj"
    }
}

Function Get-ErrorStackTrace {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ErrorStackTrace
    )
    $ErrorTraces = @()
    $StackTraceMatches = Select-String '[A-Z]+ ([^,]+), ([A-Z]:\\[^:]+|[^:]+): [A-Z]+ ([\d]+)' -Input $StackTrace -AllMatches
    If($StackTraceMatches -And $StackTraceMatches.Matches) {
        Foreach($StackTraceMatch in $StackTraceMatches.Matches) {
            $ErrorTrace = @{}
            $StackFuncName = $StackTraceMatch.Groups[1].Value 
            $ErrorTrace['StackFuncName'] = $(If($StackFuncName -ne '<ScriptBlock>') { $StackFuncName } Else { '<Script>' })
            $ScriptPath = $StackTraceMatch.Groups[2].Value
            $ScriptPathShort = ''
            $ScriptName = ''
            If(Test-Path $ScriptPath -IsValid) {
                $ScriptName = Split-Path $ScriptPath -Leaf
                $ScriptPathShort = "...\$ScriptName"
            }
            $ErrorTrace['ScriptPath'] = $ScriptPath
            $ErrorTrace['ScriptName'] = $ScriptName
            $ErrorTrace['ScriptPathShort'] = $ScriptPathShort
            $ErrorTrace['LineNum']  = $StackTraceMatch.Groups[3].Value
            $ErrorTraces += $ErrorTrace
        }
    }
    return [Array]::Reverse($ErrorTraces)
}

Function Format-ErrorPositionMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PositionMessage
    )
    $result = ''
    If($PositionMessage -match '.+\.ps1:.+\n') {
        $PosLines = $PositionMessage.Split("`n")
        If($PosLines[0] -match '^([^ ]+ )([A-Z]:\\[^:]+|[^:]+)(:.+)$') {
            $PosPrefix = $Matches[1]
            $PosPath   = $Matches[2]
            $PosEnding = $Matches[3]
            If(Test-Path $PosPath -IsValid) {
                $PosLines[0] = "$PosPrefix ...\$(Split-Path $PosPath -Leaf)$PosEnding"
                $result = $PosLines -join "`n"
            }
        }
    }
    return $result
}

Function Parse-ErrorDetails {
    [CmdletBinding()]
    Param(
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Object]$Exception
    )
    If($null -eq $ErrorRecord -And $null -eq $Exception) { return $null }
    $ErrFields = @{}
    If($Exception) {
        $ErrFields['Message']  = Get-ObjectFieldValue -FieldName 'Message'               -Object $Exception
        $ErrFields['Category'] = Get-ObjectFieldValue -FieldName 'Category'              -Object $Exception
        $ErrFields['ErrorId']  = Get-ObjectFieldValue -FieldName 'FullyQualifiedErrorId' -Object $Exception
        $ErrCodeD = Get-ObjectFieldValue -FieldName 'HResult' -Object $Exception
        $ErrFields['ErrCodeD']    = $ErrCodeD
        $ErrFields['ErrCodeH']    = $(If ($ErrCodeD) { "0x{0:X8}" -f $ErrCodeD })
        $ErrFields['Win32ErrMsg'] = $(If ($ErrCodeD) { $Msg = $(Get-ActualError $ErrCodeD).Message; $(If ($Msg -notmatch "Unknown Error \(") { $Msg }) })
        $ErrData = Get-ObjectFieldValue -FieldName 'Data' -Object $Exception
        $ErrFields['ThrowManual'] = $false
        If($ErrData) {
            $ErrFields['ThrowManual'] = Get-ObjectFieldValue -FieldName 'ThrowManual' -Object $ErrData
            $ErrFields['DebugInfo']   = Get-ObjectFieldValue -FieldName 'DebugInfo'   -Object $ErrData
            $ErrFields['CallStack']   = Get-ObjectFieldValue -FieldName 'CallStack'   -Object $ErrData
        }
    }
    If($ErrorRecord) {
        $CategoryInfo = Get-ObjectFieldValue -FieldName 'InvocationInfo' -Object $ErrorRecord
        If($CategoryInfo) {
            $ErrFields['Reason']      = Get-ObjectFieldValue -FieldName 'Reason'       -Object $CategoryInfo
            $ErrFields['TargetName']  = Get-ObjectFieldValue -FieldName 'TargetName' -Object $CategoryInfo
            $ErrFields['TargetType']  = Get-ObjectFieldValue -FieldName 'TargetType' -Object $CategoryInfo
        }
        $ErrFields['ErrDetails']  = Get-ObjectFieldValue -FieldName 'ErrorDetails' -Object $ErrorRecord
        $InvocationInfo = Get-ObjectFieldValue -FieldName 'InvocationInfo' -Object $ErrorRecord
        If($InvocationInfo) {
            $ErrFields['ErrLine']     = ([string]$(Get-ObjectFieldValue -FieldName 'Line'  -Object $InvocationInfo)).TrimEnd("`r","`n"," ")
            $ErrFields['LineNum']     = Get-ObjectFieldValue -FieldName 'ScriptLineNumber' -Object $InvocationInfo
            $ErrFields['Position']    = Get-ObjectFieldValue -FieldName 'OffsetInLine'     -Object $InvocationInfo
            $FuncName = Get-ObjectFieldValue -FieldName 'MyCommand'      -Object $InvocationInfo
            $CallName = Get-ObjectFieldValue -FieldName 'InvocationName' -Object $InvocationInfo
            $ErrFields['FuncName']    = $FuncName
            $ErrFields['CallName']    = $CallName
            $PosMsg = Get-ObjectFieldValue -FieldName 'PositionMessage' -Object $InvocationInfo
            $PosMsg = Format-ErrorPositionMessage $PosMsg
            $ErrFields['ErrPosition'] = $PosMsg
        }
        $StackTrace  = Get-ObjectFieldValue -FieldName 'ScriptStackTrace' -Object $ErrorRecord
        If($StackTrace) {
            $ErrorTraces = Get-ErrorStackTrace  $StackTrace
            $ErrFields['ErrTraces']   = $ErrorTraces
            If($null -ne $ErrorTraces -And $ErrorTraces.Count -gt 0) {
                If(-Not $FuncName) { $ErrFields['FuncName'] = $ErrorTraces[-1].StackFuncName }
            }
            $ErrorTraceMsg = $(Foreach ($ErrorTrace in $ErrorTraces) {
                "$($ErrorTrace['StackFuncName']){... at Line:$($ErrorTrace.LineNum) "
            }) -join "-> "
            $ErrFields['ErrTraceMsg'] = $ErrorTraceMsg
        }

    }
    return $ErrFields
}

Function Get-ObjectFieldValue {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [Object]$Object,
        [Parameter(Mandatory)]
        [string]$FieldName
    )
    If ($null -eq $Object) { return $null }
    If(@(Get-Member -Input $Object -Member Properties | Select -Expand Name) -contains $FieldName) {
        return $Object.$FieldName
    } ElseIf (@(Get-Member -Input $Object -Member Properties | Select -Expand Name) -contains 'Keys') {
        return $Object[$FieldName]
    }
    return $null
}

Function Traverse-ErrorRecord {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [ValidateSet('Exception','InnerException')]
        [string]$NestingType = 'Exception'
    )
    If ($NestingType -eq 'Exception') {
        If($null -ne $ErrorRecord) {
            $ExceptionObject = Get-ObjectFieldValue -FieldName 'Exception' -Object $ErrorRecord
            If($null -ne $ExceptionObject) {
                $NestedRecord = Get-ObjectFieldValue -FieldName 'ErrorRecord' -Object $ExceptionObject
                If($NestedRecord) {
                    return $NestedRecord
                }
            }
        }
    } ElseIf ($NestingType -eq 'InnerException') {
        $Global:testErr6 = $ErrorRecord
        If($null -ne $ErrorRecord) {
            $exceptionObject = Get-ObjectFieldValue -FieldName 'Exception' -Object $ErrorRecord
            If($null -ne $exceptionObject) {
                $Global:excObj2 = $exceptionObject
                $innerExceptionObject = Get-ObjectFieldValue -FieldName 'InnerException' -Object $ErrorRecord.Exception
                If($null -ne $innerExceptionObject) {
                    $innerRecord = Get-ObjectFieldValue -FieldName 'ErrorRecord' -Object $innerExceptionObject
                    If($null -ne $innerRecord) {
                        return $innerRecord
                    }
                }
            }
        }
    }
    return $null
}

Function Build-DebugErrorDetails {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $Errors = @()
    $currentRecord = $ErrorRecord
    $MaxDepth = 5
    $index = 0
    While ($currentRecord) {
        $currentError = @{}
        $exceptionObject = Get-ObjectFieldValue -Object $currentRecord -FieldName 'Exception'
        If($exceptionObject) {
            $currentError = $(Parse-ErrorDetails -ErrorRecord $currentRecord -Exception $exceptionObject)
        } Else {
            Break
        }
        $innerErrors = @()
        $innerException = Get-ObjectFieldValue -Object $exceptionObject -FieldName 'InnerException'
        If($innerException) {
            $innerRecord = Traverse-ErrorRecord -ErrorRecord $currentRecord -NestingType 'InnerException'
            If($innerRecord) { # If there is embedded ErrorRecord inside InnerException, then proceed to parse information from it
                               # Nesting structure: $currentRecord.Exception.InnerException.ErrorRecord
                $exceptionObject = Get-ObjectFieldValue -Object $innerRecord -FieldName 'Exception'
                If($exceptionObject) {
                    $innerIndex = 0
                    While ($innerRecord) {
                        $innerError = Parse-ErrorDetails -ErrorRecord $innerRecord -Exception $exceptionObject
                        $innerErrors += $innerError
                        $innerIndex++
                        If($innerIndex -gt $MaxDepth) { Break }
                        $innerRecord = Traverse-ErrorRecord -ErrorRecord $innerRecord -NestingType 'InnerException'
                    }
                }
            } Else { # Else - proceed to parse information just from InnerException
                $innerError = Parse-ErrorDetails -Exception $innerException
                $innerErrors += $innerError
                $Global:inerr1 = $innerError
            }
        }
        $currentError['innerErrors'] = $innerErrors
        $Errors += $currentError
        $index++
        If($index -gt $MaxDepth) { Break }
        $currentRecord = $(Traverse-ErrorRecord -ErrorRecord $currentRecord -NestingType Exception)
    }
    return $Errors
}

Function Format-DebugErrorMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Hashtable]$ErrorDetail
    )
    $item = $ErrorDetail
    $errorMessage = @()
    If($item.Message) { $errorMessage += "Error: $($item.Message)" }
    If($item.Reason) { $errorMessage += "Reason: $($item.Reason)" }
    If($item.Category -And $item.Category -ne 'NotSpecified') {
        $errorMessage += "Category: $($item.Category)"
    }
    If($item.ErrorId) { $errorMessage += "Error ID: $($item.ErrorId)" }
    If($item.ErrCodeD) { $errorMessage += "Error Code DEC: $($item.ErrCodeD)" }
    If($item.ErrCodeH) { $errorMessage += "Error Code HEX: $($item.ErrCodeH)" }
    If($item.Win32ErrMsg) {
        $errorMessage += "Win32 Error Message: $($item.Win32ErrMsg)"
    }
    $AliasName = $(If($item.CallName -And $item.CallName -ne $item.FuncName) { $item.CallName })
    $CalledFuncName = $(If(-Not $AliasName) { $($item.FuncName) } Else { "#$AliasName->[$CalledFuncName]" })
    If($CalledFuncName) { $errorMessage += "Occurred at: $CalledFuncName" }
    If($AliasName) {
        $errorMessage += "Invoked Alias Name: $($item.CallName)"
    }
    If($item.ThrowManual) { $errorMessage += "Thrown Manually: $($item.ThrowManual)" }
    If($item.ErrTraceMsg) { $errorMessage += "Error Trace: $($item.ErrTraceMsg)" }
    If($item.ErrDetails) { $errorMessage += "Error Details: $($item.ErrDetails)" }
    If($item.ErrLine) { $errorMessage += "Error Line: $($item.ErrLine)" }
    If($item.LineNum) { $errorMessage += "Line Number: $($item.LineNum)" }
    If($item.Position) { $errorMessage += "Char Offset: $($item.Position)" }
    If($item.ErrPosition) { $errorMessage += "Error Position: $($item.ErrPosition)" }
    If($item.TargetName) { $errorMessage += "Target Name: $($item.TargetName)" }
    If($item.TargetType) { $errorMessage += "Target Type: $($item.TargetType)" }
    If($item.ThrowManual) {
        $errorMessage += "Debug Info: $($item['DebugInfo'])"
        $errorMessage += "Debug Info: $($item['CallStack'])"
    }
    return $errorMessage
}

Function Build-DebugErrorMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Hashtable[]]$ErrorDetails
    )
    $errorMessages = @()
    Foreach ($item in $ErrorDetails) {
        $errorMessage = Format-DebugErrorMessage $item
        $errorMessages += $($errorMessage -join "`n")
        Foreach($innerItem in $item.innerErrors) {
            If($innerItem) {
                $errorMessages += "[inner Exception Error]:"
                $errorMessages += $($(Format-DebugErrorMessage $innerItem) -join "`n")
            }
        }
    }
    return $errorMessages
}

Function Compact-DuplicateErrorDetails {
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Hashtable[]]$ErrorDetails
    )
    $result = @()
    Foreach ($item in $ErrorDetails) {
        If ($result.Count -eq 0 -or $result[-1].Message -ne $item.Message) {
            # Add new entry if first item or message changed
            $result += $item.Clone()
        }
        Else {
            # Merge reasons while maintaining uniqueness
            $existingReasons = @($result[-1].Reason -split ', ') 
            $newReasons = @($item.Reason -split ', ') 
            $combined = @($existingReasons + $newReasons) | Select -Unique
            $result[-1].Reason = [string]($combined -join ', ')
        }
    }
    return $result
}

Function Get-DebugErrorMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $ErrorDetails = Build-DebugErrorDetails -ErrorRecord $ErrorRecord
    $Global:errDetails = $ErrorDetails
    $ErrorDetails = Compact-DuplicateErrorDetails $ErrorDetails
    $debugErrorMessage = $(Build-DebugErrorMessage $ErrorDetails) -join "`n"
    return $debugErrorMessage
}

Function Show-OutputExceptionInfo {
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $i = 0
    [Hashtable[]]$errors = @()
    While ($ErrorRecord.Exception) {
        $iError = @{}
        $Message = $ErrorRecord.Exception.Message
        $iError['Message'] = $Message
        $Category = $ErrorRecord.CategoryInfo.Category
        $ErrorId = $ErrorRecord.FullyQualifiedErrorId
        $Reason  = $ErrorRecord.CategoryInfo.Reason
        $ErrCodeD  = $ErrorRecord.Exception.HResult
        $ErrCodeH = "0x{0:X8}" -f $ErrCodeD
        $Win32ErrorMessage = $(Get-ActualError $ErrCodeD).Message
        $ErrorDetails = $ErrorRecord.ErrorDetails
        $ErrLine = $ErrorRecord.InvocationInfo.Line.TrimEnd("`r","`n"," ")
        $LineNum = $ErrorRecord.InvocationInfo.ScriptLineNumber
        $Position = $ErrorRecord.InvocationInfo.OffsetInLine
        $FuncName = $ErrorRecord.InvocationInfo.MyCommand
        $CallName = $ErrorRecord.InvocationInfo.InvocationName
        $TargetName = $ErrorRecord.CategoryInfo.TargetName
        $TargetType = $ErrorRecord.CategoryInfo.TargetType
        $ThrownManually = $ErrorRecord.Exception.Data['ThrownManually']
        If($ThrownManually) {
            $iError['DebugInfo'] = $ErrorRecord.Exception.Data['DebugInfo']
            $iError['CallStack'] = $ErrorRecord.Exception.Data['CallStack']
        }
        $iError['ErrorId']  = $ErrorId
        $iError['Reason']   = $Reason
        $iError['Category'] = $Category
        $iError['ErrCodeD'] = $ErrCodeD
        $iError['ErrCodeH'] = $ErrCodeH
        $iError['Win32ErrorMessage'] = $(If($Win32ErrorMessage -inotlike 'Unknown Error *') { $Win32ErrorMessage })
        $iError['ErrorDetails'] = $ErrorDetails
        $iError['ErrLine']  = $ErrLine
        $iError['LineNum']  = $LineNum
        $iError['Position'] = $Position
        $iError['TargetName'] = $TargetName
        $iError['TargetType'] = $TargetType
        $iError['FuncName'] = $FuncName
        $iError['CallName'] = $CallName
        $iError['ThrownManually'] = $ThrownManually
        $ErrorTraces = @()
        $StackTrace = $ErrorRecord.ScriptStackTrace
        $StackTraceMatches = Select-String '[A-Z]+ ([^,]+), ([A-Z]:\\[^:]+|[^:]+): [A-Z]+ ([\d]+)' -Input $StackTrace -AllMatches
        If($StackTraceMatches -And $StackTraceMatches.Matches) {
            Foreach($StackTraceMatch in $StackTraceMatches.Matches) {
                $ErrorTrace = @{}
                $StackFuncName = $StackTraceMatch.Groups[1].Value 
                $ErrorTrace['StackFuncName'] = $(If($StackFuncName -ne '<ScriptBlock>') { $StackFuncName } Else { '<Script>' })
                $ScriptPath = $StackTraceMatch.Groups[2].Value
                $ScriptPathShort = ''
                $ScriptName = ''
                If(Test-Path $ScriptPath -IsValid) {
                    $ScriptName = Split-Path $ScriptPath -Leaf
                    $ScriptPathShort = "...\$ScriptName"
                }
                $ErrorTrace['ScriptPath'] = $ScriptPath
                $ErrorTrace['ScriptName'] = $ScriptName
                $ErrorTrace['ScriptPathShort'] = $ScriptPathShort
                $ErrorTrace['LineNum']  = $StackTraceMatch.Groups[3].Value
                $ErrorTraces += $ErrorTrace
            }
        }
        $iError['ErrorTraces'] = $ErrorTraces
        $ErrorTraceMsg = ''
        [Array]::Reverse($ErrorTraces)
        If(-Not $FuncName) { $iError['FuncName']= $ErrorTraces[-1].StackFuncName }
        $ErrorTraceMsg = $(Foreach ($ErrorTrace in $ErrorTraces) {
            "$($ErrorTrace['StackFuncName']){... at Line:$($ErrorTrace.LineNum) "
        }) -join "-> "
        $iError['ErrorTraceMsg'] = $ErrorTraceMsg
        $PosMsg = $ErrorRecord.InvocationInfo.PositionMessage
        If($PosMsg -match '.+\.ps1:.+\n') {
            $PosLines = $PosMsg.Split("`n")
            If($PosLines[0] -match '^([^ ]+ )([A-Z]:\\[^:]+|[^:]+)(:.+)$') {
                $PosPrefix = $Matches[1]
                $PosPath   = $Matches[2]
                $PosEnding = $Matches[3]
                If(Test-Path $PosPath -IsValid) {
                    $PosLines[0] = "$PosPrefix ...\$(Split-Path $PosPath -Leaf)$PosEnding"
                    $PosMsg = $PosLines -join "`n"
                }
            }
        }
        $iError['ErrorPosition'] = $PosMsg
        $InnerExceptions = @()
        If($ErrorRecord.Exception.InnerException) {
            $InnerException = $ErrorRecord.Exception.InnerException
            $j = 0
            While($InnerException) {
                $iInnerException = @{}
                $iInnerException['Message'] = $InnerException.Message
                $InnerExceptions += $iInnerException
                If(@(Get-Member -Input $InnerException -Member Properties | Select -Expand Name) -contains 'InnerException') {
                    $InnerException = $InnerException.InnerException
                } Else {
                    Break
                }
                $j++
            }
        }
        $iError['InnerExceptions'] = $InnerExceptions
        $errors += $iError
        If(@(Get-Member -Input $ErrorRecord.Exception -Member Properties | Select -Expand Name) -contains 'ErrorRecord') {
            $ErrorRecord = $ErrorRecord.Exception.ErrorRecord
        } Else {
            Break
        }
        $i++
    }
    $result = @()
    Foreach ($item in $errors) {
        If ($result.Count -eq 0 -or $result[-1].Message -ne $item.Message) {
            # Add new entry if first item or message changed
            $result += $item.Clone()
        }
        Else {
            # Merge reasons while maintaining uniqueness
            $existingReasons = @($result[-1].Reason -split ', ') 
            $newReasons = @($item.Reason -split ', ') 
            $combined = @($existingReasons + $newReasons) | Select -Unique
            $result[-1].Reason = [string]($combined -join ', ')
        }
    }

    Foreach ($item in $result) {
        Write-Host "Error: $($ErrorRecord.Exception.Message)" -ForegroundColor Green
        Write-Host "Reason: $($item.Reason)" -ForegroundColor Green
        If($item.Category -And $item.Category -ne 'NotSpecified') {
            Write-Host "Category: $($item.Category)" -ForegroundColor Green
        }
        Write-Host "Error ID: $($item.ErrorId)" -ForegroundColor Green
        Write-Host "Error Code DEC: $($item.ErrCodeD)" -ForegroundColor Green
        Write-Host "Error Code HEX: $($item.ErrCodeH)" -ForegroundColor Green
        If($item.Win32ErrorMessage) {
            Write-Host "Win32 Error Message: $($item.Win32ErrorMessage)" -ForegroundColor Green
        }
        If($item.InnerExceptions) {
            Foreach($InnerException in $item.InnerExceptions) {
                Write-Host "Inner Exception Error: $($InnerException.Message)" -ForegroundColor Green
            }
        }
        $AliasName = $(If($null -ne $item.CallName -And $item.CallName -ne '' -And $item.CallName -ne $item.FuncName) { $item.CallName })
        $CalledFuncName = $(If(-Not $AliasName) { $($item.FuncName) } Else { "#$AliasName->[$CalledFuncName]" })
        Write-Host "Occurred at: $CalledFuncName" -ForegroundColor Green
        If($item.CallName) {
            Write-Host "Invoked Alias Name: $($item.CallName)" -ForegroundColor Green
        }
        Write-Host "Thrown Manually: $($item.ThrownManually)" -ForegroundColor Green
        Write-Host "Error Trace: $($item.ErrorTraceMsg)" -ForegroundColor Green
        Write-Host "Error Details: $($item.ErrorDetails)" -ForegroundColor Green
        Write-Host "Error Line: $($item.ErrLine)" -ForegroundColor Green
        Write-Host "Line Number: $($item.LineNum)" -ForegroundColor Green
        Write-Host "Char Offset: $($item.Position)" -ForegroundColor Green
        Write-Host "Position Message: $($item.ErrorPosition)" -ForegroundColor Green
        If($item.TargetName) {
            Write-Host "Target Name: $($item.TargetName)" -ForegroundColor Green
        }
        If($item.TargetType) {
            Write-Host "Target Type: $($item.TargetType)" -ForegroundColor Green
        }
        If($item.ThrownManually) {
            Write-Host "Debug Info: $($item['DebugInfo'])" -ForegroundColor Green
            Write-Host "Debug Info: $($item['CallStack'])" -ForegroundColor Green
        }
    }
}

Function Append-ThrownErrorData {
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FunctionName
    )
    $DebugInfo = $(Format-DebugInfo -Message "Error in $FunctionName" -TransitivePoint -Scope 1)
    $CallStack = $(Format-DebugInfo -Message "CallStack" -EntryPoint -Scope 1)
    $ErrorRecord.Exception.Data['DebugInfo'] = $DebugInfo
    $ErrorRecord.Exception.Data['CallStack'] = $CallStack
    $ErrorRecord.Exception.Data['ThrowManual'] = $true
    return $ErrorRecord
}


# EXAMPLE Usage:

function Get-ProcessOutput {
    param($pid)
    $proc = Get-Process -Id $pid -ErrorAction Stop
    try {
        $reader = $proc.StandardOutput
        $reader.ReadToEnd()
    } catch {
        throw $_
    }
}


Function Some-Function {
    [CmdletBinding()]
    Param(
        [bool]$ParamBool = $false,
        [switch]$ParamSwitch = $false,
        [string]$ParamString = "Other String!",
        [int]$ParamtInt = 3
    )
    $SomeParameter = 5
    $SomeOtherParam = "SomeValue1"
    function Nested-Function {
        [CmdletBinding()]
        Param(
            [bool]$TestBool,
            [switch]$TestSwitch,
            [string]$TestString = "Some String!",
            [int]$testInt = 12
        )
        $SomeParameter = 11
        $SomeOtherParam2 = "SomeValue2"

        function Process-Item {
            param([ValidateSet('1','0')]$item)
            process {

                    Get-WmiObject -Namespace "root\invalidnamespace" -Class Win32_OperatingSystem -ErrorAction Stop

                Try {

                    # 1:
                    <#
                    $divisor = 0
                    [ref]$out = $null
                    [Math]::DivRem(1,$divisor,$out)
                    #>

                    # 2:
                    <#
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Quit()
                    #>

                    # 3:
                    <#
                    $ex = [System.UnauthorizedAccessException]::new("Access denied to resource")
                    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                        $ex,
                        "AccessDenied",
                        [System.Management.Automation.ErrorCategory]::PermissionDenied,
                        $null
                    )
                    $errorRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new("Check permissions and try again.")
                    throw $errorRecord
                    #>

                    # 4:
                    #Stop-Process -Id 999999 -ErrorAction Stop

                    # 5:
                    #Get-ProcessOutput -pid 0

                    # 6:
                    #if ("test" -match "[unclosed") { }

                    # 7:
                    #$nullObject = $null
                    #$nullObject.ToString()

                    # To-Do: Fix omit duplicating innner error
                    # 8:
                    #[int]"abc"

                    # 9:
                    1,0,3 | ForEach-Object { Process-Item $_ }
                    

                } Catch {
            
                    throw $(Append-ThrownErrorData $_ -FunctionName $MyInvocation.MyCommand.Name)
                }

            }
        }

        Process-Item

    }
    Try {
        Nested-Function -TestBool $true -TestSwitch -TestString "NEw String Value!" -testInt 52 -ErrorAction Stop
    } Catch {
        $global:testErr12 = $_
        Try {
            Write-Host $(Get-DebugErrorMessage -ErrorRecord $_) -ForegroundColor Green
        } Catch {

            $global:testErr11 = $_
        }
        
    }
}


Some-Function -ErrorAction Stop

