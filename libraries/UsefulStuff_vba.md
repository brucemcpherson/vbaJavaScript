# VBA Project: **vbaJavaScript**
## VBA Module: **[UsefulStuff](/libraries/UsefulStuff.vba "source is here")**
### Type: StdModule  

This procedure list for repo (vbaJavaScript) was automatically created on 6/22/2015 2:37:54 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in UsefulStuff

---
VBA Procedure: **nameExists**  
Type: **Function**  
Returns: **name**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function nameExists(s As String) As name*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **whereIsThis**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function whereIsThis(r As Variant) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Variant|False||


---
VBA Procedure: **OpenUrl**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function OpenUrl(url) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|Variant|False||


---
VBA Procedure: **firstCell**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function firstCell(inrange As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
inrange|Range|False||


---
VBA Procedure: **lastCell**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function lastCell(inrange As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
inrange|Range|False||


---
VBA Procedure: **isSheet**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function isSheet(o As Object) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
o|Object|False||


---
VBA Procedure: **findShape**  
Type: **Function**  
Returns: **shape**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function findShape(sName As String, Optional ws As Worksheet = Nothing) As shape*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||
ws|Worksheet|True| Nothing|


---
VBA Procedure: **findRecurse**  
Type: **Function**  
Returns: **shape**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function findRecurse(target As String, co As GroupShapes) As shape*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
target|String|False||
co|GroupShapes|False||


---
VBA Procedure: **clearHyperLinks**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub clearHyperLinks(ws As Worksheet)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws|Worksheet|False||


---
VBA Procedure: **sheetExists**  
Type: **Function**  
Returns: **Worksheet**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function sheetExists(sName As String, Optional complain As Boolean = True) As Worksheet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||
complain|Boolean|True| True|


---
VBA Procedure: **wholeSheet**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function wholeSheet(wn As String) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
wn|String|False||


---
VBA Procedure: **wholeWs**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function wholeWs(ws As Worksheet) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws|Worksheet|False||


---
VBA Procedure: **wholeRange**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function wholeRange(r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **cleanFind**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function cleanFind(x As Variant, r As Range, Optional complain As Boolean = False, Optional singlecell As Boolean = False) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Variant|False||
r|Range|False||
complain|Boolean|True| False|
singlecell|Boolean|True| False|


---
VBA Procedure: **msglost**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub msglost(x As Variant, r As Range, Optional extra As String = "")*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
x|Variant|False||
r|Range|False||
extra|String|True| ""|


---
VBA Procedure: **SAd**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SAd(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rngIn|Range|False||
target|Range|True| Nothing|
singlecell|Boolean|True| False|
removeRowDollar|Boolean|True| False|
removeColDollar|Boolean|True| False|


---
VBA Procedure: **SAdOneRange**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SAdOneRange(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rngIn|Range|False||
target|Range|True| Nothing|
singlecell|Boolean|True| False|
removeRowDollar|Boolean|True| False|
removeColDollar|Boolean|True| False|


---
VBA Procedure: **AddressNoDollars**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function AddressNoDollars(a As Range, Optional doRow As Boolean = True, Optional doColumn As Boolean = True) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Range|False||
doRow|Boolean|True| True|
doColumn|Boolean|True| True|


---
VBA Procedure: **isReallyEmpty**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function isReallyEmpty(r As Range) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **toEmptyRow**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function toEmptyRow(r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **toEmptyCol**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function toEmptyCol(r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **toEmptyBox**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function toEmptyBox(r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||


---
VBA Procedure: **getLikelyColumnRange**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getLikelyColumnRange(Optional ws As Worksheet = Nothing) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws|Worksheet|True| Nothing|


---
VBA Procedure: **deleteAllFromCollection**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub deleteAllFromCollection(co As Collection)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
co|Collection|False||


---
VBA Procedure: **deleteAllShapes**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub deleteAllShapes(r As Range, startingwith As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||
startingwith|String|False||


---
VBA Procedure: **makearangeofShapes**  
Type: **Function**  
Returns: **ShapeRange**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function makearangeofShapes(r As Range, startingwith As String) As ShapeRange*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
r|Range|False||
startingwith|String|False||


---
VBA Procedure: **UTF16To8**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function UTF16To8(ByVal UTF16 As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||


---
VBA Procedure: **URLEncode**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function URLEncode( StringVal As String, Optional SpaceAsPlus As Boolean = False, Optional UTF8Encode As Boolean = True ) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
StringVal|String|False||
SpaceAsPlus|Boolean|True| False|
UTF8Encode|Boolean|True| True|


---
VBA Procedure: **cloneFormat**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub cloneFormat(b As Range, a As Range)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
b|Range|False||
a|Range|False||


---
VBA Procedure: **SortColl**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function SortColl(ByRef coll As Collection, eorder As Long) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByRef|Collection|False||
eorder|Long|False||


---
VBA Procedure: **getHandle**  
Type: **Function**  
Returns: **Integer**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getHandle(sName As String, Optional readOnly As Boolean = False) As Integer*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||
readOnly|Boolean|True| False|


---
VBA Procedure: **afConcat**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function afConcat(arr() As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
arr|Variant|False||


---
VBA Procedure: **quote**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function quote(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **q**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function q() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **qs**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function qs() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **bracket**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function bracket(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **list**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function list(ParamArray args() As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ParamArray|Variant|False||


---
VBA Procedure: **qlist**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function qlist(ParamArray args() As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ParamArray|Variant|False||


---
VBA Procedure: **diminishingReturn**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function diminishingReturn(val As Double, Optional s As Double = 10) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
val|Double|False||
s|Double|True| 10|


---
VBA Procedure: **pivotCacheRefreshAll**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Sub pivotCacheRefreshAll()*  

**no arguments required for this procedure**


---
VBA Procedure: **makeKey**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function makeKey(v As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
v|Variant|False||


---
VBA Procedure: **Base64Encode**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Base64Encode(sText)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sText|Variant|False||


---
VBA Procedure: **Stream_StringToBinary**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Stream_StringToBinary(Text)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Text|Variant|False||


---
VBA Procedure: **Stream_BinaryToString**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Stream_BinaryToString(Binary)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
Binary|Variant|False||


---
VBA Procedure: **Base64Decode**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function Base64Decode(ByVal base64String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|Variant|False||


---
VBA Procedure: **openNewHtml**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function openNewHtml(sName As String, sContent As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||
sContent|String|False||


---
VBA Procedure: **readFromFile**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function readFromFile(sName As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sName|String|False||


---
VBA Procedure: **arrayLength**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function arrayLength(a) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Variant|False||


---
VBA Procedure: **getControlValue**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getControlValue(ctl As Object) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ctl|Object|False||


---
VBA Procedure: **setControlValue**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function setControlValue(ctl As Object, v As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ctl|Object|False||
v|Variant|False||


---
VBA Procedure: **isinCollection**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function isinCollection(vCollect As Variant, sid As Variant) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
vCollect|Variant|False||
sid|Variant|False||


---
VBA Procedure: **getLatFromDistance**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getLatFromDistance(mLat As Double, d As Double, heading As Double) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
mLat|Double|False||
d|Double|False||
heading|Double|False||


---
VBA Procedure: **getLonFromDistance**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getLonFromDistance(mLat As Double, mLon As Double, d As Double, heading As Double) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
mLat|Double|False||
mLon|Double|False||
d|Double|False||
heading|Double|False||


---
VBA Procedure: **earthRadius**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function earthRadius() As Double*  

**no arguments required for this procedure**


---
VBA Procedure: **toRadians**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function toRadians(deg)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
deg|Variant|False||


---
VBA Procedure: **fromRadians**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function fromRadians(rad) As Double*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rad|Variant|False||


---
VBA Procedure: **dimensionCount**  
Type: **Function**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function dimensionCount(a As Variant) As Long*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
a|Variant|False||


---
VBA Procedure: **min**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function min(ParamArray args() As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ParamArray|Variant|False||


---
VBA Procedure: **max**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function max(ParamArray args() As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ParamArray|Variant|False||


---
VBA Procedure: **encloseTag**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function encloseTag(tag As String, Optional newLine As Boolean = True, Optional tClass As String = vbNullString, Optional args As Variant) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
tag|String|False||
newLine|Boolean|True| True|
tClass|String|True| vbNullString|
args|Variant|True||


---
VBA Procedure: **scrollHack**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function scrollHack() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **escapeify**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function escapeify(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **unEscapify**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function unEscapify(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **basicStyle**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function basicStyle() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **tableStyle**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function tableStyle() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **is64BitExcel**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function is64BitExcel() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **includeJQuery**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function includeJQuery() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **includeGoogleCallBack**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function includeGoogleCallBack(c As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
c|String|False||


---
VBA Procedure: **jScriptTag**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function jScriptTag(Optional src As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
src|String|True||


---
VBA Procedure: **jDivAtMouse**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function jDivAtMouse()*  

**no arguments required for this procedure**


---
VBA Procedure: **toClipBoard**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function toClipBoard(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **importTabbed**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function importTabbed(fn As String, r As Range) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
fn|String|False||
r|Range|False||


---
VBA Procedure: **biasedRandom**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function biasedRandom(possibilities, weights) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
possibilities|Variant|False||
weights|Variant|False||


---
VBA Procedure: **sleep**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub sleep(seconds As Long)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
seconds|Long|False||


---
VBA Procedure: **getDateFromTimestamp**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getDateFromTimestamp(s As String) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **dateFromUnix**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function dateFromUnix(s As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|Variant|False||


---
VBA Procedure: **isSomething**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function isSomething(o As Object) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
o|Object|False||


---
VBA Procedure: **tinyTime**  
Type: **Function**  
Returns: **Double**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function tinyTime() As Double*  

**no arguments required for this procedure**


---
VBA Procedure: **getTableRange**  
Type: **Function**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getTableRange(tableName As String, Optional complain As Boolean = True) As Range*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
tableName|String|False||
complain|Boolean|True| True|


---
VBA Procedure: **getListObject**  
Type: **Function**  
Returns: **ListObject**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getListObject(tableName As String) As ListObject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
tableName|String|False||


---
VBA Procedure: **listObjectExists**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function listObjectExists(ws As Worksheet, sName As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ws|Worksheet|False||
sName|String|False||
