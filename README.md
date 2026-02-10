# vba-rules
Semgrep basic rules for Visual Basic for Applications (VBA)

# Steps to add the rules to your semgrep instance
```
git clone https://github.com/r2c-CSE/vba-rules
cd vba-rules
semgrep login
semgrep publish .
```

You will see a message like this:
```
sebasrevuelta@Sebastians-Laptop vba-rules % semgrep publish .
Found 10 configs to publish with visibility "VisibilityState.ORG_PRIVATE"
--> Uploading vba-auto-exec.yaml (test cases: ['vba-auto-exec.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-auto-exec
--> Uploading vba-native-api-declaration.yaml (test cases: ['vba-native-api-declaration.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-native-api-declaration
--> Uploading vba-file-drop.yaml (test cases: ['vba-file-drop.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-file-drop
--> Uploading vba-environment-discovery.yaml (test cases: ['vba-environment-discovery.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-environment-discovery
--> Uploading vba-shell-execution.yaml (test cases: ['vba-shell-execution.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-shell-execution
--> Uploading vba-doc-var-hiding.yaml (test cases: ['vba-doc-var-hiding.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-doc-var-hiding
--> Uploading vba-hidden-window.yaml (test cases: ['vba-hidden-window.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-hidden-window
--> Uploading vba-adodb-binary-write.yaml (test cases: ['vba-adodb-binary-write.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-adodb-binary-write
--> Uploading vba-downloader-objects.yaml (test cases: ['vba-downloader-objects.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-downloader-objects
--> Uploading vba-string-reversal-obfuscation.yaml (test cases: ['vba-string-reversal-obfuscation.vba'])
    Published VisibilityState.ORG_PRIVATE rule at https://semgrep.dev/r/dash5ast.vba-string-reversal-obfuscation
```
It will add all the rules to your personal registry, but they will not be used in the scan until you add them to whatever policy/mode:
* Using API you can call this endpoint: https://semgrep.dev/api/v1/deployments/{deploymentId}/policies/{policyId}
* Using the UI, you can add rules to your policies by clicking the "Add to Policy" option in the "Rules & Policies" -> "Editor" menu:

![Adding rules to policy](images/add-policy.png)

