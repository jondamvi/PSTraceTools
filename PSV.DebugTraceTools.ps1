<#
.SYNOPSIS
    Title   : PSV.DebugTraceTools.ps1
    Author  : Jon Damvi
    Version : 1.0.1
    Date    : 02.06.2025
    License : MIT (LICENSE)

   Release Notes:
        v1.0.0 (19.05.2025) - initial release (by Jon Damvi).
        v1.0.1 (02.06.2025) - synopsis function documentation (by Jon Damvi).

.DESCRIPTION
    Shows detailed Exception Error Trace information

.INPUTS
    None.

.OUTPUTS
    None.

#>

Class ActualError {
    [int]$HResult
    [string]$HResultHex
    [string]$Facility
    [int]$FacilityCode
    [int]$ErrorCode
    [bool]$IsFailure
    [string]$Message
}

<#
.SYNOPSIS
	Title   : Function Get-ActualError
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	Translates decimal exception HResult error code to actual error information.

.PARAMETER hresult
	(Optional) Specifies HResult error code to decode.
	Expected type: [int]

.INPUTS
	None

.OUTPUTS
	ActualError - Class-object containing detailed error information.

.EXAMPLE
	[ActualError]$ErrorInfo = Get-ActualError -hresult -2146233087

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
function Get-ActualError {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [int]$HResult
    )
    # Define a class to hold the decoded error info
    Class ActualError {
        [int]$HResult
        [string]$HResultHex
        [string]$Facility
        [int]$FacilityCode
        [int]$ErrorCode
        [bool]$IsFailure
        [string]$Message
    }
    $ActualError = [ActualError]::new()
    $ActualError.HResult = $HResult
    $ActualError.HResultHex = "0x" + $HResult.ToString("X8")
    # Extract error code (lowest 16 bits)
    $ActualError.ErrorCode = $HResult -band 0xFFFF
    # Extract facility code (bits 16-30)
    $ActualError.FacilityCode = ($HResult -shr 16) -band 0x1FFF
    # Check if severity bit (bit 31) is set (failure)
    $ActualError.IsFailure = ($HResult -band 0x80000000) -ne 0
    # Get the Win32 error message corresponding to the HRESULT
    Try {
        $ActualError.Message = (New-Object System.ComponentModel.Win32Exception($HResult)).Message
    } Catch {
        $ActualError.Message = "Unknown error message"
    }
    # Map facility codes to names based on standard HRESULT facility codes
    switch ($ActualError.FacilityCode) {
        0   { $ActualError.Facility = 'FACILITY_NULL' }
        1   { $ActualError.Facility = 'FACILITY_RPC' }
        2   { $ActualError.Facility = 'FACILITY_DISPATCH' }
        3   { $ActualError.Facility = 'FACILITY_STORAGE' }
        4   { $ActualError.Facility = 'FACILITY_ITF' }
        5   { $ActualError.Facility = 'FACILITY_WIN32' }
        6   { $ActualError.Facility = 'FACILITY_WINDOWS' }
        7   { $ActualError.Facility = 'FACILITY_SSPI' }
        default { $ActualError.Facility = 'UNKNOWN' }
    }
    return $ActualError
}

<#
.SYNOPSIS
	Title   : Function Get-ActiveScopeCount
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	Gets count of active powershell scopes excluding console and global session scope.

.INPUTS
	None

.OUTPUTS
	System.Byte - count of active powershell scopes.

.EXAMPLE
    Get-ActiveScopeCount 

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Get-ActiveScopeCount {
    [CmdletBinding()]
    Param()
    [byte]$Count = 0
    While ($true) {
        Try {
            # Try to get reserved variable at scope $count. This function call limits processing overhead to minimum as possible.
            Get-Variable -Name 'PID' -ValueOnly -Scope $Count -ErrorAction Stop -WarningAction SilentlyContinue -InformationAction SilentlyContinue -Verbose:$false | Out-Null
            $Count++
        }
        Catch {
            break
        }
    }
    return [byte]($Count-2)
}

<#
.SYNOPSIS
	Title   : Function Get-LocalSetVariables
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER Scope
	(Mandatory) Specifies ...
	Expected type  : [Byte]
	Allowed Values : [0;255]
	Default Value  : 0

.INPUTS
	None

.OUTPUTS
	System.Management.Automation.PSVariable[]

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Get-LocalSetVariables -Scope $Scope

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Get-LocalSetVariables {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateRange(0,255)]
        [byte]$Scope = 0
    )
    # Scope = 0 - value is for the user
    # ActualScope is one level back from 0
    [byte]$ActualScope = $Scope + 1
    If($ActualScope -ge $(Get-ActiveScopeCount -ErrorAction Stop)) {
        return
    }
    # Known automatic variables to exclude
    [string[]]$FilterVars = @(
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
    [int]$ExcludeFlags = [int](2 -bor 8)
    # Get local variables excluding Constant and AllScope
    [System.Management.Automation.PSVariable[]]$LocalVars = [System.Management.Automation.PSVariable[]]@(
        Get-Variable -Scope $ActualScope | ? { ([int]$_.Options -band $ExcludeFlags) -eq 0 }
    )
    # Exclude known automatic variables by name
    [System.Management.Automation.PSVariable[]]$FilteredVars = [System.Management.Automation.PSVariable[]]@(
        $LocalVars.GetEnumerator() | ? { $FilterVars -notcontains $_.Name }
    )
    [System.Management.Automation.PSVariable[]]$result = [System.Management.Automation.PSVariable[]]@(
        $FilteredVars.GetEnumerator() | ? {
            [System.Management.Automation.PSVariable]$ParentVar = [System.Management.Automation.PSVariable](
                Get-Variable -Name $_.Name -Scope Script -ErrorAction SilentlyContinue
            )
            If ($null -eq $ParentVar) {
                $true
            }
            Else {
                $_.Value -ne $ParentVar.Value
            }
        }
    )
    return [System.Management.Automation.PSVariable[]]$result
}

<#
.SYNOPSIS
	Title   : Function Format-TraceInfo
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER Message
	(Mandatory) Specifies ...
	Expected type: [String]

.PARAMETER EntryPoint
	(Mandatory) Specifies ...
	Expected type: [SwitchParameter]Aliases: Entry

.PARAMETER TransitivePoint
	(Mandatory) Specifies ...
	Expected type: [SwitchParameter]Aliases: Transitive

.PARAMETER Scope
	(Optional) Specifies ...
	Expected type: [Byte]
	Allowed Values: [0;255]
	Default Value: 0

.INPUTS
	None

.OUTPUTS
	$TraceMessage - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Format-TraceInfo -Message $Message -EntryPoint $EntryPoint -TransitivePoint $TransitivePoint -Scope $Scope

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Format-TraceInfo {
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
        [byte]$ActualScope = $Scope + 1
        If($ActualScope -ge $(Get-ActiveScopeCount -ErrorAction Stop)) {
            return
        }
        # Function body as before, using $PSCmdlet.ParameterSetName to branch logic
        $CallStack = @(Get-PSCallStack -ErrorAction Stop)
        $CallStackScope = $ActualScope
        If($CallStack.Count -le $CallStackScope) {
            return
        }
        $CurrentFunction = $CallStack[$CallStackScope].FunctionName
        $InvocationName  = $CallStack[$CallStackScope].InvocationInfo.InvocationName
        $AliasUsed = ($InvocationName -And $InvocationName -ne $CurrentFunction)
        $EntryPointName = "${CurrentFunction}"
        If ($AliasUsed) {
            $EntryPointName = "#$InvocationName->${CurrentFunction}"
        }
        # Build call stack string only if in Params set
        $CallStackStr = ''
        If ($PSCmdlet.ParameterSetName -eq 'EntryPoint') {
            $FilteredStack = @($CallStack[$CallStackScope..($CallStack.Count-1)] | % {
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
            [Array]::Reverse($FilteredStack)
            $CallStackStr = $FilteredStack -join ':'
        }
        # Format items (Params or Locals)
        $Params = $CallStack[$CallStackScope].InvocationInfo.BoundParameters
        Switch ($PSCmdlet.ParameterSetName) {
            'EntryPoint' {
                $items = $Params
            }
            'TransitivePoint' {
                $items = @{}
                $VariableScope  = $ActualScope
                $LocalVariables = @(
                    Get-LocalSetVariables -Scope ($VariableScope) -ErrorAction Stop
                ) | ? { 
                    (-Not $Params.ContainsKey($_.Name) -OR $Params[$_.Name] -ne $_.Value)
                }
                Foreach ($var in $LocalVariables) {
                    $items[$var.Name] = $var.Value
                }
            }
            default {
                $items = @{}
            }
        }
        # Format items using your Format-Value helper
        $ItemStrings = @(Foreach ($key in $items.Keys | Sort) {
            $Value = $items[$key]
            "${key} = " + (Format-Value $Value -ErrorAction Stop)
        })
        $ItemsStr = If ($ItemStrings.Count) { $ItemStrings -join ", `n " } Else { '' }
        # Determine line number if call stack available
        $LineNum = $CallStack[$CallStackScope].ScriptLineNumber
        If ($PSCmdlet.ParameterSetName -eq 'EntryPoint') {
            $ItemsStr = "(`n $ItemsStr`n)"
        } Else {
            $ItemsStr = "`n $ItemsStr`n"
        }
        $TraceMessage = ''
        If ($PSCmdlet.ParameterSetName -eq 'EntryPoint') {
            $TraceMessage = "[${EntryPointName}]: '${Message}' (Line:${LineNum}): ${CallStackStr}${ItemsStr}"
        } Else {
            $TraceMessage = "[${EntryPointName}]: '${Message}' (Line:${LineNum}): `n{${CallStackStr}${ItemsStr}}"
        }
        return $TraceMessage
    }
    Catch {
        Throw $_
    }
}

<#
.SYNOPSIS
	Title   : Function Format-Value
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER Value
	(Optional) Specifies ...
	Expected type: [Object]

.INPUTS
	None

.OUTPUTS
	[string] - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Format-Value -Value $Value

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Format-Value {
    [CmdletBinding()]
    Param(
        [Object]$Value
    )
    $MaxLength = 255
    [string]$ReturnStr = ''
    If ($null -eq $Value) {
        $ReturnStr = '`$null'
    }
    ElseIf (($Value -is [switch]) -OR ($Value -is [bool])) {
        $ReturnStr = $Value.ToString().ToLower()
    }
    ElseIf ($Value -is [string] -OR $Value -is [ScriptBlock]) {
        If ($Value -is [ScriptBlock]) { $Value = $Value.ToString() }
        $Str = $Value.TrimStart(" ", "`n", "`r")
        # Get up to first newline or $MaxLength, whichever is shorter
        $NewlineIdx = $Str.IndexOf("`n")
        If ($NewlineIdx -ge 0 -And $NewlineIdx -lt $MaxLength) {
            $Preview = $Str.Substring(0, $NewlineIdx).TrimEnd(" ", "`r", "`n")
            $ReturnStr = '"' + $Preview + ' \\~..."'
        } ElseIf ($str.Length -gt $MaxLength) {
            $Preview = $Str.Substring(0, $MaxLength).TrimEnd(" ", "`r", "`n")
            $ReturnStr = '"' + $Preview + ' ~..."'
        } Else {
            $ReturnStr = '"' + $Str + '"'
        }
    }
    ElseIf ($Value.GetType().IsPrimitive) {
        return $Value.ToString()
    }
    ElseIf ($Value -is [System.Array]) {
        $TypeName = $Value.GetType().GetElementType().Name
        $Count = $Value.Count
        $ReturnStr = "[$TypeName[$Count]]`$Obj"
    }
    Else {
        Try {
            $ValueString = $Value.ToString()
            $ValueTypeFullName = $Value.GetType().FullName
            If ($ValueString -ne $ValueTypeFullName) {
                $ReturnStr = "'" + $ValueString + "'"
            }
        } Catch {

        }
        $ValueTypeName = $Value.GetType().Name
        $ReturnStr = "[$ValueTypeName]`$Obj"
    }
    return $ReturnStr
}

<#
.SYNOPSIS
	Title   : Function Get-ErrorStackTrace
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorStackTrace
	(Mandatory) Specifies ...
	Expected type: [String]

.INPUTS
	None

.OUTPUTS
	None

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Get-ErrorStackTrace -ErrorStackTrace $ErrorStackTrace

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Get-ErrorStackTrace {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ErrorStackTrace
    )
    [Hashtable[]]$ErrorTraces = @()
    $StackTraceMatches = Select-String '[A-Z]+ ([^,]+), ([A-Z]:\\[^:]+|[^:]+): [A-Z]+ ([\d]+)' -Input $StackTrace -AllMatches
    If($StackTraceMatches -And $StackTraceMatches.Matches) {
        Foreach($StackTraceMatch in $StackTraceMatches.Matches) {
            $ErrorTrace = @{}
            $StackFuncName = $StackTraceMatch.Groups[1].Value 
            $ErrorTrace['StackFuncName'] = $(If ($StackFuncName -ne '<ScriptBlock>') { $StackFuncName } Else { '<Script>' })
            $ScriptPath = $StackTraceMatch.Groups[2].Value
            $ScriptPathShort = ''
            $ScriptName = ''
            If (Test-Path $ScriptPath -IsValid) {
                $ScriptName = Split-Path $ScriptPath -Leaf -ErrorAction Stop
                $ScriptPathShort = "...\$ScriptName"
            }
            $ErrorTrace['ScriptPath'] = $ScriptPath
            $ErrorTrace['ScriptName'] = $ScriptName
            $ErrorTrace['ScriptPathShort'] = $ScriptPathShort
            $ErrorTrace['LineNum']  = $StackTraceMatch.Groups[3].Value
            $ErrorTraces += $ErrorTrace
        }
    }
    [Hashtable[]]::Reverse($ErrorTraces)
    return $ErrorTraces
}

<#
.SYNOPSIS
	Title   : Function Format-ErrorPositionMessage
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER PositionMessage
	(Mandatory) Specifies ...
	Expected type: [String]

.INPUTS
	None

.OUTPUTS
	$result - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Format-ErrorPositionMessage -PositionMessage $PositionMessage

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Format-ErrorPositionMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$PositionMessage
    )
    [string]$result = ''
    If ($PositionMessage -match '.+\.ps1:.+\n') {
        $PosLines = $PositionMessage.Split("`n")
        If ($PosLines[0] -match '^([^ ]+ )([A-Z]:\\[^:]+|[^:]+)(:.+)$') {
            $PosPrefix = $Matches[1]
            $PosPath   = $Matches[2]
            $PosEnding = $Matches[3]
            If(Test-Path $PosPath -IsValid) {
                $PosLines[0] = "$PosPrefix ...\$(Split-Path $PosPath -Leaf -ErrorAction Stop)$PosEnding"
                $result = $PosLines -join "`n"
            }
        }
    }
    return [string]$result
}

<#
.SYNOPSIS
	Title   : Function Parse-ErrorDetails
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorRecord
	(Optional) Specifies ...
	Expected type: [ErrorRecord]

.PARAMETER Exception
	(Mandatory) Specifies ...
	Expected type: [Object]

.INPUTS
	None

.OUTPUTS
	$ErrFields - returned when ...
	$null      - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Parse-ErrorDetails -ErrorRecord $ErrorRecord -Exception $Exception

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Parse-ErrorDetails {
    [CmdletBinding()]
    Param(
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Object]$Exception
    )
    If ($null -eq $ErrorRecord -And $null -eq $Exception) { return $null }
    [Hashtable]$ErrFields = @{}
    If($Exception) {
        $ErrFields['Message']  = Get-ObjectFieldValue -FieldName 'Message'               -Object $Exception -ErrorAction Stop


        $ErrCodeD = Get-ObjectFieldValue -FieldName 'HResult' -Object $Exception -ErrorAction Stop
        $ErrFields['ErrCodeD']    = $ErrCodeD
        $ErrFields['ErrCodeH']    = $(If ($ErrCodeD) { "0x{0:X8}" -f $ErrCodeD })
        $ErrFields['Win32ErrMsg'] = $(If ($ErrCodeD) { $Msg = $(Get-ActualError $ErrCodeD -ErrorAction Stop).Message; $(If ($Msg -notmatch "Unknown Error \(") { $Msg }) })
        $ErrData = Get-ObjectFieldValue -FieldName 'Data' -Object $Exception -ErrorAction Stop
        $ErrFields['ThrowManual'] = $false
        If($ErrData) {
            $ErrFields['ThrowManual'] = Get-ObjectFieldValue -FieldName 'ThrowManual' -Object $ErrData -ErrorAction Stop
            $ErrFields['TraceInfo']   = Get-ObjectFieldValue -FieldName 'TraceInfo'   -Object $ErrData -ErrorAction Stop
            $ErrFields['CallStack']   = Get-ObjectFieldValue -FieldName 'CallStack'   -Object $ErrData -ErrorAction Stop
        }
    }
    If($ErrorRecord) {
        $ErrFields['ErrorId']  = Get-ObjectFieldValue -FieldName 'FullyQualifiedErrorId' -Object $ErrorRecord -ErrorAction Stop
        $CategoryInfo = Get-ObjectFieldValue -FieldName 'CategoryInfo' -Object $ErrorRecord
        If($CategoryInfo) {
            $ErrFields['Category']    = Get-ObjectFieldValue -FieldName 'Category'   -Object $CategoryInfo -ErrorAction Stop
            $ErrFields['Reason']      = Get-ObjectFieldValue -FieldName 'Reason'     -Object $CategoryInfo -ErrorAction Stop
            $ErrFields['TargetName']  = Get-ObjectFieldValue -FieldName 'TargetName' -Object $CategoryInfo -ErrorAction Stop
            $ErrFields['TargetType']  = Get-ObjectFieldValue -FieldName 'TargetType' -Object $CategoryInfo -ErrorAction Stop
        }
        $ErrFields['ErrDetails']  = Get-ObjectFieldValue -FieldName 'ErrorDetails' -Object $ErrorRecord -ErrorAction Stop
        $InvocationInfo = Get-ObjectFieldValue -FieldName 'InvocationInfo' -Object $ErrorRecord -ErrorAction Stop
        If($InvocationInfo) {
            $ErrFields['ErrLine']     = ([string]$(Get-ObjectFieldValue -FieldName 'Line'  -Object $InvocationInfo -ErrorAction Stop)).TrimEnd("`r","`n"," ")
            $ErrFields['LineNum']     = Get-ObjectFieldValue -FieldName 'ScriptLineNumber' -Object $InvocationInfo -ErrorAction Stop
            $ErrFields['Position']    = Get-ObjectFieldValue -FieldName 'OffsetInLine'     -Object $InvocationInfo -ErrorAction Stop
            $FuncName = Get-ObjectFieldValue -FieldName 'MyCommand'      -Object $InvocationInfo -ErrorAction Stop
            $CallName = Get-ObjectFieldValue -FieldName 'InvocationName' -Object $InvocationInfo -ErrorAction Stop
            $ErrFields['FuncName']    = $FuncName
            $ErrFields['CallName']    = $CallName
            $PosMsg = Get-ObjectFieldValue -FieldName 'PositionMessage' -Object $InvocationInfo -ErrorAction Stop
            $PosMsg = Format-ErrorPositionMessage $PosMsg -ErrorAction Stop
            $ErrFields['ErrPosition'] = $PosMsg
        }
        $StackTrace  = Get-ObjectFieldValue -FieldName 'ScriptStackTrace' -Object $ErrorRecord -ErrorAction Stop
        If($StackTrace) {
            [Hashtable[]]$ErrorTraces = [Hashtable[]](Get-ErrorStackTrace  $StackTrace -ErrorAction Stop)
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
    return [Hashtable]$ErrFields
}

<#
.SYNOPSIS
	Title   : Function Get-ObjectFieldValue
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER Object
	(Mandatory) Specifies ...
	Expected type: [Object]

.PARAMETER FieldName
	(Mandatory) Specifies ...
	Expected type: [String]

.INPUTS
	None

.OUTPUTS
	$null - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Get-ObjectFieldValue -Object $Object -FieldName $FieldName

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Get-ObjectFieldValue {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [Object]$Object,

        [Parameter(Mandatory)]
        [string]$FieldName
    )
    If ($null -eq $Object) { return $null }
    If(@(Get-Member -Input $Object -Member Properties -ErrorAction Stop | Select -Expand Name) -contains $FieldName) {
        return $Object.$FieldName
    } ElseIf (@(Get-Member -Input $Object -Member Properties -ErrorAction Stop | Select -Expand Name) -contains 'Keys') {
        return $Object[$FieldName]
    }
    return $null
}

<#
.SYNOPSIS
	Title   : Function Traverse-ErrorRecord
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorRecord
	(Mandatory) Specifies ...
	Expected type: [ErrorRecord]

.PARAMETER NestingType
	(Mandatory) Specifies ...
	Expected type: [String]
	Allowed Values: Exception, InnerException
	Default Value: 'Exception'

.INPUTS
	None

.OUTPUTS
	$NestedRecord - returned when ...
	$innerRecord  - returned when ...
	$null         - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Traverse-ErrorRecord -ErrorRecord $ErrorRecord -NestingType $NestingType

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
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
            $ExceptionObject = Get-ObjectFieldValue -FieldName 'Exception' -Object $ErrorRecord -ErrorAction Stop
            If($null -ne $ExceptionObject) {
                $NestedRecord = Get-ObjectFieldValue -FieldName 'ErrorRecord' -Object $ExceptionObject -ErrorAction Stop
                If($NestedRecord) {
                    return $NestedRecord
                }
            }
        }
    } ElseIf ($NestingType -eq 'InnerException') {
        If($null -ne $ErrorRecord) {
            $exceptionObject = Get-ObjectFieldValue -FieldName 'Exception' -Object $ErrorRecord -ErrorAction Stop
            If($null -ne $exceptionObject) {
                $innerExceptionObject = Get-ObjectFieldValue -FieldName 'InnerException' -Object $ErrorRecord.Exception -ErrorAction Stop
                If($null -ne $innerExceptionObject) {
                    $innerRecord = Get-ObjectFieldValue -FieldName 'ErrorRecord' -Object $innerExceptionObject -ErrorAction Stop
                    If($null -ne $innerRecord) {
                        return $innerRecord
                    }
                }
            }
        }
    }
    return $null
}

<#
.SYNOPSIS
	Title   : Function Build-TraceErrorDetails
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorRecord
	(Mandatory) Specifies ...
	Expected type: [ErrorRecord]

.INPUTS
	None

.OUTPUTS
	$Errors - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Build-TraceErrorDetails -ErrorRecord $ErrorRecord

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Build-TraceErrorDetails {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    [Hashtable[]]$Errors = @()
    $currentRecord = $ErrorRecord
    [byte]$MaxDepth = 5
    [byte]$index = 0
    While ($currentRecord) {
        $currentError = @{}
        $exceptionObject = Get-ObjectFieldValue -Object $currentRecord -FieldName 'Exception' -ErrorAction Stop
        If ($exceptionObject) {
            $currentError = $(Parse-ErrorDetails -ErrorRecord $currentRecord -Exception $exceptionObject -ErrorAction Stop)
        } Else {
            break
        }
        $innerErrors = @()
        $innerException = Get-ObjectFieldValue -Object $exceptionObject -FieldName 'InnerException' -ErrorAction Stop
        If ($innerException) {
            $innerRecord = Traverse-ErrorRecord -ErrorRecord $currentRecord -NestingType 'InnerException' -ErrorAction Stop
            If ($innerRecord) { # If there is embedded ErrorRecord inside InnerException, then proceed to parse information from it
                               # Nesting structure: $currentRecord.Exception.InnerException.ErrorRecord
                $exceptionObject = Get-ObjectFieldValue -Object $innerRecord -FieldName 'Exception' -ErrorAction Stop
                If ($exceptionObject) {
                    $innerIndex = 0
                    While ($innerRecord) {
                        $innerError = Parse-ErrorDetails -ErrorRecord $innerRecord -Exception $exceptionObject -ErrorAction Stop
                        $innerErrors += $innerError
                        $innerIndex++
                        If ($innerIndex -gt $MaxDepth) { break }
                        $innerRecord = Traverse-ErrorRecord -ErrorRecord $innerRecord -NestingType 'InnerException' -ErrorAction Stop
                    }
                }
            } Else { # Else - proceed to parse information just from InnerException
                $innerError = Parse-ErrorDetails -Exception $innerException -ErrorAction Stop
                $innerErrors += $innerError
            }
        }
        $currentError['innerErrors'] = $innerErrors
        $Errors += $currentError
        $index++
        If ($index -gt $MaxDepth) { break }
        $currentRecord = $(Traverse-ErrorRecord -ErrorRecord $currentRecord -NestingType Exception -ErrorAction Stop)
    }
    return $Errors
}

<#
.SYNOPSIS
	Title   : Function Format-TraceErrorMessage
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorDetail
	(Mandatory) Specifies ...
	Expected type: [Hashtable]

.INPUTS
	None

.OUTPUTS
	$errorMessage - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Format-TraceErrorMessage -ErrorDetail $ErrorDetail

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Format-TraceErrorMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Hashtable]$ErrorDetail
    )
    [Hashtable]$item = $ErrorDetail
    [string[]]$errorMessage = [string[]]@()
    [string]$ErrorDelimiterLine = [string]$('-' * 80)
    If($item.Message) { $errorMessage += "$($item.Message)" }
    $errorMessage += $ErrorDelimiterLine
    $errorMessage += '[TRACE Error Details]'

    If($item.Reason)  { $errorMessage += "Reason: $($item.Reason)" }
    If($item.Category -And $item.Category -ne 'NotSpecified') {
        $errorMessage += "Category: $($item.Category)"
    }
    If($item.ErrorId)  { $errorMessage += "Error ID: $($item.ErrorId)" }
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
        $errorMessage += $ErrorDelimiterLine
        $errorMessage += "[TRACE Variables]: $($item['TraceInfo'])"
        $errorMessage += $ErrorDelimiterLine
        $errorMessage += "[TRACE CallStack]: $($item['CallStack'])"
        $errorMessage += $ErrorDelimiterLine
    }
    return [string[]]$errorMessage
}

<#
.SYNOPSIS
	Title   : Function Build-TraceErrorMessage
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorDetails
	(Mandatory) Specifies ...
	Expected type: [Hashtable[]]

.INPUTS
	None

.OUTPUTS
	$errorMessages - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Build-TraceErrorMessage -ErrorDetails $ErrorDetails

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Build-TraceErrorMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Hashtable[]]$ErrorDetails
    )
    [string[]]$errorMessages = [string[]]@()
    [string]$ErrorDelimiterLine = [string]$('-' * 80)
    Foreach ($item in $ErrorDetails) {
        [string[]]$errorMessage = [string[]]$(Format-TraceErrorMessage $item -ErrorAction Stop)
        $errorMessages += $($errorMessage -join "`n")
        Foreach($innerItem in $item.innerErrors) {
            If($innerItem) {
                $errorMessages += "[Inner Exception Error]:"
                $errorMessages += $([string[]]$(Format-TraceErrorMessage $innerItem -ErrorAction Stop) -join "`n")
                $errorMessages += $ErrorDelimiterLine
            }
        }
    }
    return [string[]]$errorMessages
}

<#
.SYNOPSIS
	Title   : Function Compact-DuplicateErrorDetails
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorDetails
	(Mandatory) Specifies ...
	Expected type: [Hashtable[]]

.INPUTS
	None

.OUTPUTS
	$result - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Compact-DuplicateErrorDetails -ErrorDetails $ErrorDetails

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Compact-DuplicateErrorDetails {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Hashtable[]]$ErrorDetails
    )
    [Hashtable[]]$result = [Hashtable[]]@()
    Foreach ($item in $ErrorDetails) {
        If ($result.Count -eq 0 -OR $result[-1].Message -ne $item.Message) {
            # Add new entry if first item or message changed
            $result += $item.Clone()
        } Else {
            # Merge reasons while maintaining uniqueness
            [string[]]$existingReasons = [string[]]@($result[-1].Reason -split ', ') 
            [string[]]$newReasons = [string[]]@($item.Reason -split ', ') 
            [string[]]$combined = [string[]]@(@($existingReasons + $newReasons) | Select -Unique)
            $result[-1].Reason = [string]($combined -join ', ')
        }
    }
    return [Hashtable[]]$result
}

<#
.SYNOPSIS
	Title   : Function Get-TraceErrorMessage
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorRecord
	(Mandatory) Specifies ...
	Expected type: [ErrorRecord]

.INPUTS
	None

.OUTPUTS
	$TraceErrorMessage - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Get-TraceErrorMessage -ErrorRecord $ErrorRecord

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Get-TraceErrorMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $ErrorDetails = Build-TraceErrorDetails -ErrorRecord $ErrorRecord -ErrorAction Stop
    $ErrorDetails = Compact-DuplicateErrorDetails $ErrorDetails -ErrorAction Stop
    [string]$TraceErrorMessage = [string]($(Build-TraceErrorMessage $ErrorDetails -ErrorAction Stop) -join "`n")
    return [string]$TraceErrorMessage
}

<#
.SYNOPSIS
	Title   : Function Append-ThrownErrorData
	Author  : Jon Damvi
	Version : 1.0.0
	Date    : 01.06.2025
	License : MIT (LICENSE)

	Release Notes: 
		v1.0.0 (01.06.2025) - initial release (by Jon Damvi).

.DESCRIPTION
	[Description to be added]

.PARAMETER ErrorRecord
	(Mandatory) Specifies ...
	Expected type: [ErrorRecord]

.PARAMETER FunctionName
	(Mandatory) Specifies ...
	Expected type: [String]

.INPUTS
	None

.OUTPUTS
	$ErrorRecord - returned when ...

.EXAMPLE
	# Usage Case ... Example Description :
	PS > Append-ThrownErrorData -ErrorRecord $ErrorRecord -FunctionName $FunctionName

.LINK
	https://github.com/jondamvi/PSV.DebugTraceTools
#>
Function Append-ThrownErrorData {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FunctionName
    )
    If (-Not $FunctionName) {
        [string]$FunctionName = [string](@(Get-PSCallStack -ErrorAction Stop)[1].FunctionName)
    }
    [string]$TraceInfo = [string]$(Format-TraceInfo -Message "Error in $FunctionName" -TransitivePoint -Scope 1 -ErrorAction Stop)
    [string]$CallStack = [string]$(Format-TraceInfo -Message "CallStack" -EntryPoint -Scope 1 -ErrorAction Stop)
    $ErrorRecord.Exception.Data['TraceInfo'] = $TraceInfo
    $ErrorRecord.Exception.Data['CallStack'] = $CallStack
    $ErrorRecord.Exception.Data['ThrowManual'] = [bool]$true
    return [System.Management.Automation.ErrorRecord]$ErrorRecord
}

Function Confirm-ErrorThrowManual {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    If ($null -ne $ErrorRecord.Exception.Data) {
        return [bool]$ErrorRecord.Exception.Data['ThrowManual']
    }
    return [bool]$false
}
# -------------------------------------------------------------------------------------------------
# EXAMPLE Usage:

function Get-ProcessOutput {
    param($pid)
    $proc = Get-Process -Id $pid -ErrorAction Stop
    try {
        $reader = $proc.StandardOutput
        $reader.ReadToEnd()
    } catch {
        Throw $_
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
                    # [ERROR SIMULATION - TEST CASES]:

                    # 0: Test WMI Query Invalid Namespace:
                    #Get-WmiObject -Namespace "root\invalidnamespace" -Class Win32_OperatingSystem -ErrorAction Stop

                Try {
                    
                    # 1: Test Division by Zero (.NET operation):
                    $divisor = 0
                    [ref]$out = $null
                    [Math]::DivRem(1,$divisor,$out)

                    # 2: Test Uknown Object Type:
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Quit()

                    # 3: Test Custom Access Denied Error:
                    $ex = [System.UnauthorizedAccessException]::new("Access denied to resource")
                    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                        $ex,
                        "AccessDenied",
                        [System.Management.Automation.ErrorCategory]::PermissionDenied,
                        $null
                    )
                    $errorRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new("Check permissions and try again.")
                    throw $errorRecord

                    # 4: Test Non-existent process stop:
                    #Stop-Process -Id 999999 -ErrorAction Stop

                    # 5: Test Non-existent process query:
                    #Get-ProcessOutput -pid 0

                    # 6: Test Invalid Regex Pattern:
                    #if ("test" -match "[unclosed") { }

                    # 7: Test Method invocation on null-object:
                    #$nullObject = $null
                    #$nullObject.ToString()

                    # To-Do: Fix omit duplicating innner error
                    # 8: Test incorrect type value assignment:
                    #[int]"abc"

                    # 9: Test Pipeline Error:
                    #1,0,3 | ForEach-Object { Process-Item $_ }

                    # 10: Test Division by Zero (Powershell Operation):
                    #1/0
                } Catch {
                    Write-Host $(Get-TraceErrorMessage -ErrorRecord $_ -ErrorAction Stop) -ForegroundColor Red
                    Write-Host ""
                    Write-Host "--------------------------------------------------------------------"
                    Write-Host ""
                    Throw $(Append-ThrownErrorData $_ -FunctionName $MyInvocation.MyCommand.Name -ErrorAction Stop)
                }
            }
        }
        Process-Item
    }
    Try {
        Nested-Function -TestBool $true -TestSwitch -TestString "NEw String Value!" -testInt 52 -ErrorAction Stop
    } Catch {
        # Observe here Rich-Exception Data Trace Details: Trace Variables, Trace CallStack and Indication of manually thrown exception:
        Write-Host $(Get-TraceErrorMessage -ErrorRecord $_ -ErrorAction Stop) -ForegroundColor Magenta
    }
}

Some-Function -ErrorAction Stop



<#
# Error Simulation Advanced Test-Cases:
# Example 1: Creating an ErrorRecord that contains another ErrorRecord
function New-NestedErrorRecordStructure {
    # Create the innermost ErrorRecord
    $innermostException = New-Object System.InvalidOperationException("Core failure")
    $innermostErrorRecord = New-Object System.Management.Automation.ErrorRecord(
        $innermostException,
        "CoreError",
        [System.Management.Automation.ErrorCategory]::InvalidOperation,
        "InnerTarget"
    )
    
    # Create an exception that wraps the ErrorRecord
    $wrapperException = New-Object System.Management.Automation.RuntimeException(
        "Middle layer error",
        $null,  # No inner exception here
        $innermostErrorRecord  # But we pass the ErrorRecord as additional data
    )
    
    # Add the ErrorRecord to the exception's data
    $wrapperException.Data["ErrorRecord"] = $innermostErrorRecord
    
    # Create a middle ErrorRecord
    $middleErrorRecord = New-Object System.Management.Automation.ErrorRecord(
        $wrapperException,
        "MiddleError",
        [System.Management.Automation.ErrorCategory]::OperationStopped,
        "MiddleTarget"
    )
    
    # Create the outermost exception that contains the middle ErrorRecord
    $outermostException = New-Object System.Management.Automation.RemoteException(
        "Outer layer error",
        $null,
        $middleErrorRecord
    )
    
    # Set the ErrorRecord property on the exception
    $outermostException.GetType().GetProperty(
        "ErrorRecord", 
        [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Instance
    ).SetValue($outermostException, $middleErrorRecord)
    
    # Create the final ErrorRecord
    $outerErrorRecord = New-Object System.Management.Automation.ErrorRecord(
        $outermostException,
        "OuterError",
        [System.Management.Automation.ErrorCategory]::RemoteOperationError,
        "OuterTarget"
    )
    
    return $outerErrorRecord
}

# Example 2: Simulating remote/job error nesting
function Test-RemoteErrorNesting {
    [CmdletBinding()]
    Param()
    [ScriptBlock]$scriptBlock = [ScriptBlock]{
        try {
            Throw "Initial error in remote session"
        }
        catch {
            # This creates the first ErrorRecord
            Write-Error "$($_.Exception)"
        }
    }
    # Run in a job to simulate remote execution
    $job = Start-Job -ScriptBlock $scriptBlock -ErrorAction Continue
    Wait-Job $job | Out-Null
    
    try {
        # Receive-Job wraps errors in additional ErrorRecord layers
        Receive-Job $job -ErrorAction Stop
    }
    catch {
        # This error will have nested ErrorRecord structure
        return $_
    }
    finally {
        Remove-Job $job -Force
    }
}

# Example 3: Manually creating the nested structure you mentioned
# Corrected Example 3: Creating deeply nested ErrorRecord structures
function New-DeepNestedErrorRecord {
    # Level 1 - Innermost
    $level1Exception = New-Object System.Exception("Level 1 error")
    $level1ErrorRecord = New-Object System.Management.Automation.ErrorRecord(
        $level1Exception, "Error1", "InvalidOperation", $null
    )
    
    # Level 2 - Wrap Level 1 ErrorRecord in a custom exception
    $level2Exception = New-Object System.Exception("Level 2 error")
    # Add the ErrorRecord as a property
    Add-Member -InputObject $level2Exception -MemberType NoteProperty -Name "ErrorRecord" -Value $level1ErrorRecord -Force
    
    $level2ErrorRecord = New-Object System.Management.Automation.ErrorRecord(
        $level2Exception, "Error2", "InvalidOperation", $null
    )
    
    # Level 3 - Contains Level 2 ErrorRecord
    $level3Exception = New-Object System.Exception("Level 3 error")
    Add-Member -InputObject $level3Exception -MemberType NoteProperty -Name "ErrorRecord" -Value $level2ErrorRecord -Force
    
    $level3ErrorRecord = New-Object System.Management.Automation.ErrorRecord(
        $level3Exception, "Error3", "InvalidOperation", $null
    )
    
    # Level 4 - Outermost
    $level4Exception = New-Object System.Exception("Level 4 error")
    Add-Member -InputObject $level4Exception -MemberType NoteProperty -Name "ErrorRecord" -Value $level3ErrorRecord -Force
    
    $level4ErrorRecord = New-Object System.Management.Automation.ErrorRecord(
        $level4Exception, "Error4", "InvalidOperation", $null
    )
    
    return $level4ErrorRecord
}


# Corrected Real-world test case 2: Custom ErrorRecord subclass with proper property naming
Add-Type @'
using System;
using System.Management.Automation;

namespace CustomErrors {
    public class ExtendedErrorRecord : ErrorRecord {
        public ErrorRecord ErrorRecord { get; set; }  // Named exactly as your framework expects
        public string RemoteComputerName { get; set; }
        
        public ExtendedErrorRecord(Exception exception, string errorId, 
            ErrorCategory errorCategory, object targetObject, ErrorRecord errorRecord) 
            : base(exception, errorId, errorCategory, targetObject) {
            this.ErrorRecord = errorRecord;
        }
    }
    
    public class RemotingException : Exception {
        public ErrorRecord ErrorRecord { get; set; }  // Also follows the naming convention
        public string OriginInfo { get; set; }
        
        public RemotingException(string message, ErrorRecord errorRecord) : base(message) {
            this.ErrorRecord = errorRecord;
        }
    }
    
    // Additional complex exception type that might be seen in production
    public class NestedRemotingException : Exception {
        public ErrorRecord ErrorRecord { get; set; }
        new public Exception InnerException { get; set; }
        
        public NestedRemotingException(string message, ErrorRecord errorRecord, Exception inner) 
            : base(message, inner) {
            this.ErrorRecord = errorRecord;
            this.InnerException = inner;
        }
    }
}
'@

function Test-CustomErrorRecordSubclass {
    [CmdletBinding()]
    Param()
    
    # Level 1: Base error
    $level1Exception = New-Object System.IO.FileNotFoundException(
        "Cannot find file on remote system"
    )
    
    $level1ErrorRecord = New-Object System.Management.Automation.ErrorRecord(
        $level1Exception,
        "RemoteFileNotFound",
        [System.Management.Automation.ErrorCategory]::ObjectNotFound,
        "\\RemoteServer\Share\file.txt"
    )
    
    # Level 2: Wrap in custom remoting exception
    $remotingException = New-Object CustomErrors.RemotingException(
        "Remote operation failed",
        $level1ErrorRecord
    )
    $remotingException.OriginInfo = "RemoteServer01"
    
    # Level 3: Create extended error record with ErrorRecord property
    $extendedErrorRecord = New-Object CustomErrors.ExtendedErrorRecord(
        $remotingException,
        "RemoteOperationFailed", 
        [System.Management.Automation.ErrorCategory]::ProtocolError,
        $null,
        $level1ErrorRecord  # This will be accessible via .ErrorRecord property
    )
    $extendedErrorRecord.RemoteComputerName = "RemoteServer01"
    
    return $extendedErrorRecord
}

# Even more complex scenario with multiple ErrorRecord properties at different levels
function Test-DeepCustomNesting {
    [CmdletBinding()]
    Param()
    
    # Create a chain where multiple objects have ErrorRecord properties
    
    # Bottom level
    $bottomEx = New-Object System.Exception("Bottom level error")
    $bottomError = New-Object System.Management.Automation.ErrorRecord(
        $bottomEx, "BottomError", "InvalidOperation", $null
    )
    
    # Middle level - Exception with ErrorRecord property
    $middleEx = New-Object System.Exception("Middle level error")
    Add-Member -InputObject $middleEx -MemberType NoteProperty -Name "ErrorRecord" -Value $bottomError -Force
    
    $middleError = New-Object System.Management.Automation.ErrorRecord(
        $middleEx, "MiddleError", "InvalidOperation", $null
    )
    
    # Top level - Custom exception with both ErrorRecord and nested structure
    $topEx = New-Object CustomErrors.NestedRemotingException(
        "Top level error",
        $middleError,  # This goes to ErrorRecord property
        $middleEx      # This goes to InnerException
    )
    
    # Create a custom ErrorRecord subclass as the final wrapper
    $finalError = New-Object CustomErrors.ExtendedErrorRecord(
        $topEx,
        "ComplexNestedError",
        [System.Management.Automation.ErrorCategory]::SecurityError,
        $null,
        $middleError  # Another ErrorRecord reference
    )
    
    return $finalError
}

# Usage examples testing advanced cases:

#Write-Host "`nTraversing the structure:" -ForegroundColor Green
#Get-NestedErrorRecordInfo -ErrorRecord $nestedError

#Write-Host "`nTesting with real remote execution:" -ForegroundColor Green
#Test-InvokeCommandNesting

Function Test-ErrorSimulationMain {
    [CmdletBinding()]
    Param()
    Function Test-ErrorSimulationSubfunctionAAA {
        [CmdletBinding()]
        Param([byte]$Index,[string]$Name,[switch]$Enable)
            Function Test-ErrorSimulationSubfunctionBBB {
                [CmdletBinding()]
                Param([string]$Value,[switch]$Execute,[byte]$Count)

                #Write-Host "Creating nested ErrorRecord structure..." -ForegroundColor Green
                #$nestedError = New-DeepNestedErrorRecord
                #Throw $nestedError

                #$RemoteJobError = Test-RemoteErrorNesting -ErrorAction Stop
                #Throw $RemoteJobError

                #$TestCase1 = New-NestedErrorRecordStructure
                #Throw $TestCase1

                #$CustomErrorRecordSubclass = Test-CustomErrorRecordSubclass
                #Throw $CustomErrorRecordSubclass

                #$CustomEvenMoreComplexCase = Test-DeepCustomNesting
                #Throw $CustomEvenMoreComplexCase
            }
            Try {
                Test-ErrorSimulationSubfunctionBBB -Execute -Value "ThirdScope" -Count 7 -ErrorAction Stop
            } Catch {
                Write-Host $(Get-TraceErrorMessage -ErrorRecord $_ -ErrorAction Stop) -ForegroundColor Red
                Write-Host ""
                Write-Host "--------------------------------------------------------------------"
                Write-Host ""
                Throw $(Append-ThrownErrorData $_ -FunctionName $MyInvocation.MyCommand.Name -ErrorAction Stop)
            }
    }
    Try {
        Test-ErrorSimulationSubfunctionAAA -Index 3 -Name "SecondScope" -Enable -ErrorAction Stop
    } Catch {
        # Observe here Rich-Exception Data Trace Details: Trace Variables, Trace CallStack and Indication of manually thrown exception:
        Write-Host $(Get-TraceErrorMessage -ErrorRecord $_ -ErrorAction Stop) -ForegroundColor Magenta
    }
}

Test-ErrorSimulationMain -ErrorAction Stop

#>
