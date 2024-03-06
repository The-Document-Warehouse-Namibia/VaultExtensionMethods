# M-Files Vault Extension Method

This repository contains a custom M-Files Vault extension method that allows you to retrieve IDs by alias for object types, classes, and property definitions.

## Overview

The extension method is designed to parse input strings containing a delimiter (;) and execute the appropriate function based on the provided function name. It supports three main functions:

- GetObjectTypeIDByAlias: Retrieves the object type ID by alias.
- GetClassIDByAlias: Retrieves the class ID by alias.
- GetPropertyDefIDByAlias: Retrieves the property definition ID by alias.

## Usage

Prior to utilizing this method, ensure that you implement it in a Vault Extension Method. To use the method you can simply just call the method from any of your scripts like the example below.

```vbscript
Dim output
output = Vault.ExtensionMethodOperations.ExecuteVaultExtensionMethod("[VaultExtensionMethodName]", "[Alias]" & ";" & "[MethodName]")
```

## Implementation Guide

For detailed implementation guidance and examples, please refer to the M-Files Developer Portal.

## Exception Handling

Custom exception handling is implemented to handle scenarios where the specified object type, class, or property definition alias does not exist or where there are multiple objects with the same alias.

## Example

```vbscript
Option Explicit

' Define aliases.
Const OBJECT_TYPE_ALIAS = "[Object type alias]"
Const CLASS_ALIAS = "[Class alias]"
Const PROERDEF_ALIAS = "[Property def alias]"

' Resolve aliases.
Public ObjTypeID : ObjTypeID = GetIDByAlias("GetObjectTypeIDByAlias", OBJECT_TYPE_ALIAS)

Public ClassID : ClassID = GetIDByAlias("GetClassIDByAlias", CLASS_ALIAS)

Public PropID : PropID = GetIDByAlias("GetPropertyDefIDByAlias", PROERDEF_ALIAS)


Err.Raise MFScriptCancel, "Object Type ID: " & ObjTypeID & vbCrLf & _
							"Property ID: " & PropID & vbCrLf & _
							"Class ID: " & ClassID


''' HELPER FUNCTIONS

Private Function GetIDByAlias(alias, method)
	' Vault extension method requires a ";" delimiter to separate the alias and method name.
	GetIDByAlias = Vault.ExtensionMethodOperations.ExecuteVaultExtensionMethod("TDW.VaultExtensionMethod.GetIDByAlias", _
			 alias & ";" & method)
End Function
```
