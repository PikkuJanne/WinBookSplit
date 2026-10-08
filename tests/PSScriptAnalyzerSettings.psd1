@{
    # Gate new harness code with defect/security rules. The unchanged application
    # is analyzed with the default rule set and recorded separately as legacy.
    IncludeRules = @(
        'PSAvoidUsingInvokeExpression',
        'PSAvoidGlobalVars',
        'PSAvoidUsingCmdletAliases',
        'PSAvoidUsingEmptyCatchBlock',
        'PSAvoidUsingPlainTextForPassword',
        'PSAvoidUsingConvertToSecureStringWithPlainText',
        'PSAvoidUsingUsernameAndPasswordParams',
        'PSUseDeclaredVarsMoreThanAssignments',
        'PSUsePSCredentialType',
        'PSUseCompatibleSyntax'
    )
    Rules = @{
        PSUseCompatibleSyntax = @{ Enable = $true; TargetVersions = @('5.1', '7.0') }
    }
}
