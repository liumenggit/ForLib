;作者：刘老六
;脚本：未命名
;博客：http://xueahk.com
;语言：AutoHotkey1.1.24
;时间：2016年10月10日21:12:30
#Persistent	
#SingleInstance force
#Include %A_ScriptDir%\JSON.ahk
#Include *i %A_ScriptDir%\Lib\Habit\Habit.ahk
#Include *i %A_ScriptDir%\Lib\WIN\WIN.ahk
Path_data=%A_ScriptDir%\Program_list.mdb
IfNotExist,%Path_data%	;判断数据库文件是否存在 不存在则创建数据库
{
	Catalog:=ComObjCreate("ADOX.Catalog")
	Catalog.Create("Provider='Microsoft.Jet.OLEDB.4.0';Data Source=" Path_data)
	Sql_Run("CREATE TABLE globals(Pname varchar(255),Proute varchar(255),Pfun varchar(255))")	;添加全局数据库表
	Sql_Run("CREATE TABLE Program_list(Pname varchar(255),Proute varchar(255),Pfun Memo)")	;添加程序数据库表
	Sql_Run("CREATE TABLE Fun_list(funname varchar(255),fundescribe varchar(255),funcategory varchar(255),funEdition varchar(255),funswitch varchar(255))")
}
Menu,程序列表2,Add,添加分组,Menu_AddGroup
Menu,程序列表2,Add,删除程序(&D),Menu_DellProgram

Menu,程序列表3,Add,删除函数(&D),Menu_DellFun
Menu,FunMenuList,Add,查看详细(&V),Menu_Fun_Info
Menu,FunMenuList,Add,下载函数(&D),Menu_Fun_Download
;预设测试
;预备添加
/*
IsLabel(LabelName)：
*/
global ImageListID
global get
global Gv_TrreProgram
ImageListID := IL_Create(10)
IL_Add(ImageListID,"shell32.dll",246)
IL_Add(ImageListID,"shell32.dll",246)
Gui +LastFound
Gui,Add,TreeView,xm ym r20  w300 vTreeView gTreeView ImageList%ImageListID% AltSubmit
Gui,Add,ListView,y335 xm w300 r5 vDeployList glist Grid -Multi -ReadOnly  NoSortHdr,参数|说明
Gui,Add,Tab2,ym w600 h20 Buttons Section Bottom
MyImageList := IL_Create()
SendMessage, 0x1303, 0, MyImageList, SysTabControl321
IL_Add(MyImageList,"ico\download.ico")
AddTab(1,"函数商店")
IL_Add(MyImageList, "shell32.dll","270")
AddTab(2,"本地函数")
IL_Add(MyImageList,"ico\Sql_runner.ico")
AddTab(3,"数据信息")
Gui,Tab,1
DownList:=DownList()
Gui,Add,DropDownList,vListName gSwitchClass Choose1 xs Section,%DownList%
Gui,Add,Edit,vContent ys w400
Gui,Add,Button,ys hp+ gSelect Default,搜索
Gui,Add,ListView,h385 w600 xs vFunListView gFunListView AltSubmit,名称|描述|版本|更新时间|分类|热度|状态
Gui,Tab,3
Gui,Add,DropDownList,xs ys Choose1 Section,%DownList%
Gui,Add,ComboBox, ys w300
Gui,Add,Button,ys hp+  ,执行语句
Gui,Add,ListView,h385 w600 xs
Gui,Add,StatusBar,, Bar's starting text (omit to start off empty).
SB_SetText("There are " . RowCount . " rows selected.")
Gui,Show
Gui,+HwndMyGuiHwnd
Get_FunList("AllList","","")
Load_TreeList("SELECT * FROM Program_list ORDER BY Pfun DESC") 	;从数据库中依次读取程序的函数信息加载TerrView
loop{	;这是一个主循环这里添加窗口信息
	WinGet,WinGet_ID,ID,A
	WinGet,WinGet_ProcessName,ProcessName,A
	WinGet,WinGet_ProcessPath,ProcessPath,A
	StringReplace,WinGet_ProcessName,WinGet_ProcessName,.exe,,All
	if Sql_Get("SELECT COUNT(*) FROM Program_list WHERE Pname='" WinGet_ProcessName "'"){	;检查数据库中是否有本程序的数据库列 程序名不为空 程序图标部为空
		StartFun(WinGet_ProcessName,"WinWait")
	}else if GetIconCount(WinGet_ProcessPath){	;数据库中不存在本程序在数据库中添加此程序信息
		Sql_Run("Insert INTO Program_list (Pname,Proute) VALUES ('" WinGet_ProcessName "','" WinGet_ProcessPath "')")
	}
	WinWaitNotActive,ahk_id %WinGet_ID%
	StartFun(WinGet_ProcessName,"WinNotActive")
	Load_TreeList("SELECT * FROM Program_list ORDER BY Pfun DESC") 
}
return

StartFun(WinGet_ProcessName,Start){ ;激活函数
	Arr_GetProcess := JSON.load(Sql_Get("SELECT Pfun FROM Program_list WHERE Pname='" WinGet_ProcessName "'"))
	For K,V in Arr_GetProcess
	{
		IniRead,Starts,%A_ScriptDir%\Lib\%K%\%K%.ini,SectionName,Start
		if (Start=Starts){
			;#Include *i e:\OneDrive\AHK\Lib\Habit\Habit.ahk
			%K%(Arr_GetProcess[K])
		}
	}
}
Menu_Fun_Info:	;查看函数详细信息
	RunFunUrl:=get[Fun_EventInfo].FunUrl
	Run,%RunFunUrl%
return
Menu_Fun_Download:	;下载函数
	FunDownload:=get[Fun_EventInfo].FunDownload

	SplitPath,FunDownload,DownloadName
	DownloadUrl=%A_ScriptDir%\Download\%DownloadName%
	MsgBox % FunDownload
	UrlDownloadToFile,%FunDownload%,%A_ScriptDir%\Download\%DownloadName%
	if ErrorLevel{
		MsgBox,下载出错
		return
	}
	DownloadLib=%A_ScriptDir%\Lib
	SmartZip(DownloadUrl,DownloadLib)
	MsgBox,%DownloadName%下载完成
return

SwitchClass:	;分类按钮
	Gui,Submit,NoHide
	Get_FunList("Select","FunClass",ListName)
return
DownList(){	;向服务器请求所有分类名称
	static req := ComObjCreate("Msxml2.XMLHTTP")
	req.open("GET","http://139.196.173.237/autohotkey.php?FunName=DownList",false)
	req.Send()
	return req.responseText
}
FunListView:	;单击函数显示详细些菜单
if (A_GuiEvent="RightClick"){
	Fun_EventInfo:=A_EventInfo
	Menu,FunMenuList,Show
}
return

/*
GuiContextMenu:
TV_GetText(OutputVar,TV_GetSelection())
;MsgBox % OutputVar
return
*/
Select:
	Gui,Submit,NoHide
	Get_FunList("Select","FunName",Content)
return
TreeView:
TV_GetText(OutputVar,A_EventInfo)
TV_Modify(A_EventInfo,"Select")
if (A_GuiEvent="RightClick"){	;对项目右键时取得当前项的最顶级菜单加上当前项所在的等级1234 根据这两个信息去显示制定的Menu
	Menu_Structure:=TreeList(A_EventInfo)[TreeList(A_EventInfo).Length()] "" TreeList(A_EventInfo).Length()
	Try {
		Menu,Submenu1,Delete
		Loop,%A_ScriptDir%\Lib\*.*,2,1
			Menu,Submenu1,Add,%A_LoopFileName%,Menu_AddFun
		Menu,程序列表2,Add,添加功能(&A),:Submenu1
	}catch e{
		Loop,%A_ScriptDir%\Lib\*.*,2,1
			Menu,Submenu1,Add,%A_LoopFileName%,Menu_AddFun
		Try Menu,程序列表2,Add,添加功能(&A),:Submenu1
	}
	Try Menu,%Menu_Structure%,Show
	catch e
		return
	return
}
if (A_GuiEvent="DoubleClick"){	;双击项的时候获取此项是否为最子项目 从数据库总获取此子项目中的配置信息如{"配置1":"1111"}将此信息展现在ListView
	if not TV_GetChild(A_EventInfo)
		TV_FatherID:=A_EventInfo
	TV_Structure :=
	while TV_FatherID{
		TV_GetText(TV_DoubleName,TV_FatherID)
		TV_FatherID:=TV_GetParent(TV_FatherID)
		if (TV_FatherID=Gv_TrreProgram){	;到达程序功能层
			Arr_TvDouble := JSON.load(Sql_Get("SELECT Pfun FROM Program_list WHERE Pname='" TV_DoubleName "'"))
			GuiControl,-Redraw,TreeView
			LV_Delete()
			For K,V in Get_TvChild(Arr_TvDouble,TV_Structure)
				LV_Add("Icon" . set_ico,V,K)
			LV_ModifyCol()
			LV_ModifyCol(2,"AutoHdr")
			GuiControl,+Redraw,TreeView
			Break
		}
		TV_Structure:=TV_DoubleName "," TV_Structure
	}
}
return

TreeList(TV_FatherID){
	TreeList:=Object()
	while TV_FatherID{
		TV_GetText(TV_CurrentName,TV_FatherID)
		TreeList[A_Index]:=TV_CurrentName
		TV_FatherID:=TV_GetParent(TV_FatherID)
	}
	return TreeList
}
return

Get_FunList(FunName,ListName,Content){
	Gui,ListView,FunListView
	LV_Delete()
	GuiControl,-Redraw,FunListView
	static req := ComObjCreate("Msxml2.XMLHTTP")
	req.open("GET","http://139.196.173.237/autohotkey.php?FunName=" FunName "&ListName=" ListName "&Content=" Content "",false)
	req.Send()
	get:=JSON.load(req.responseText)
	For k,v in get{
		FunName:=get[A_Index].FunName
		IfExist,%A_ScriptDir%\lib\%FunName%
			FunPlace:="本地"
		else
			FunPlace:="网络"
			FunInf:=get[A_Index].FunInf
			FunTime:=get[A_Index].FunTime
			StringLeft,FunInf,FunInf,30
			FormatTime,FunTime,FunTime,yyyy/MM/dd
		LV_Add("" , get[A_Index].FunName,FunInf,get[A_Index].FunVersion,FunTime,get[A_Index].FunClass,get[A_Index].FunHeat,FunPlace)
	}
	LV_ModifyCol()
	GuiControl,+Redraw,FunListView
	Gui,ListView,DeployList
	return 
}

Menu_AddGroup:	;添加分组
	MsgBox,添加分组 %TV_CurrentName%
return

Menu_DellProgram:	;删除程序
	TV_GetText(TV_CurrentName,TV_GetSelection())
	Sql_Get("DELETE FROM Program_list WHERE Pname='" TV_CurrentName "'")
	Load_TreeList("SELECT * FROM Program_list ORDER BY Pfun DESC") 
return

Menu_DellFun:	;删除函数
	TV_GetText(TV_FunName,TV_GetSelection())
	TV_GetText(TV_ProgramName,TV_GetParent(TV_GetSelection()))
	Arr_TvDouble := JSON.load(Sql_Get("SELECT Pfun FROM Program_list WHERE Pname='" TV_ProgramName "'"))
	Arr_TvDouble.Delete(TV_FunName)
	Sql_Run("UPDATE Program_list SET Pfun ='" JSON.Dump(Arr_TvDouble) "' WHERE Pname='" TV_ProgramName "'")
	Load_TreeList("SELECT * FROM Program_list ORDER BY Pfun DESC") 
return

Menu_AddFun:	;添加函数 在程序功能的子程序中添加函数 获取当前项目的JSON字符串并添加对象到JSON存储回数据库
	TV_GetText(TV_CurrentName,TV_GetSelection())
	Arr_TvDouble := JSON.load(Sql_Get("SELECT Pfun FROM Program_list WHERE Pname='" TV_CurrentName "'"))
	if not IsObject(Arr_TvDouble){
		Add_FunName:=A_ThisMenuItem
		Arr_TvDouble := Object()
	}else if IsObject(Arr_TvDouble[A_ThisMenuItem]){
		InputBox,Add_FunName,重命名函数,请保留函数名字%A_ThisMenuItem%，在后方加入描述如%A_ThisMenuItem%-功能。,,,,,,,,%A_ThisMenuItem%
		if ErrorLevel
			return
		;goto,Menu_AddFun
	}else{
		Add_FunName:=A_ThisMenuItem
	}
	IniRead,FunINI,%A_ScriptDir%\Lib\%A_ThisMenuItem%\%A_ThisMenuItem%.ini,SectionName,Array
	Arr_TvDouble[Add_FunName]:=JSON.load(FunINI)
	Sql_Run("UPDATE Program_list SET Pfun ='" JSON.Dump(Arr_TvDouble) "' WHERE Pname='" TV_CurrentName "'")
	Load_TreeList("SELECT * FROM Program_list ORDER BY Pfun DESC") 
return

List:	;在ListView中单击第一列开始编辑 编辑完毕后保存信息到数据库
if (A_GuiEvent="e"){
	LV_GetText(LV_KEY,A_EventInfo,2)
	LV_GetText(LV_Var,A_EventInfo,1)
	Get_TvChild(Arr_TvDouble,TV_Structure)[LV_KEY]:=LV_Var
	Sql_Run("UPDATE Program_list SET Pfun ='" JSON.Dump(Arr_TvDouble) "' WHERE Pname='" TV_DoubleName "'")
}
LV_ModifyCol()
LV_ModifyCol(2,"AutoHdr")
return

Sql_Run(SQL){	;向数据库运行命令
	Recordset := ComObjCreate("ADODB.Recordset")
	Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . A_ScriptDir . "\Program_list.mdb")
	return
}
Sql_Get(SQL){	;向数据库运行命令请求返回
	Recordset := ComObjCreate("ADODB.Recordset")
	Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . A_ScriptDir . "\Program_list.mdb")
	Try return Recordset.Fields[0].Value
	catch e
		return
}
Load_TreeList(SQL){
	Recordset := ComObjCreate("ADODB.Recordset")
	Recordset.Open(SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" . A_ScriptDir . "\Program_list.mdb")
	IfWinExist,ahk_id %MyGuiHwnd%
		return
	GuiControl,-Redraw,TreeView
	TV_Delete()
	Gv_TrreProgram:=TV_Add("程序列表",,"Expand")
	while !Recordset.EOF
	{
		Proute:=Recordset.Fields["Proute"].Value
		IfExist,%Proute%
		{
			TV_FatherID:=TV_Add(Recordset.Fields["Pname"].Value,Gv_TrreProgram,"Icon"IL_Add(ImageListID,Proute,1))
			;MsgBox % Recordset.Fields["Pfun"].Value
			Arr_ProgramFun := JSON.load(Recordset.Fields["Pfun"].Value)
			Load_TvChild(Arr_ProgramFun,TV_FatherID)
		}else{
			Sql_Get("DELETE FROM Program_list WHERE Proute='" Proute "'")
		}
		Recordset.MoveNext()
	}
	GuiControl,+Redraw,TreeView
	return
}

SmartZip(s, o, t = 4)
{
    IfNotExist, %s%
        return, -1
    oShell := ComObjCreate("Shell.Application")
    if InStr(FileExist(o), "D") or (!FileExist(o) and (SubStr(s, -3) = ".zip"))
    {
        if !o
            o := A_ScriptDir
        else ifNotExist, %o%
                FileCreateDir, %o%
        Loop, %o%, 1
            sObjectLongName := A_LoopFileLongPath
        oObject := oShell.NameSpace(sObjectLongName)
        Loop, %s%, 1
        {
            oSource := oShell.NameSpace(A_LoopFileLongPath)
            oObject.CopyHere(oSource.Items, t)
        }
    }
}

Get_TvChild(Arr,Structure){	;取得TerrView程序子项目在对象中的结构信息如Array[1][2][3]
	Structure := RTrim(Structure,",")
	Loop,Parse,Structure,`,
		Arr:=Arr[A_LoopField]
	return Arr
}

Load_TvChild(Arr_ProgramFun,TV_FatherID){	;解析多层数组展示为TreeView树状图
	For K,V in Arr_ProgramFun{
		if (IsObject(V)){
			TV_FatherID:=TV_Add(K,TV_FatherID,"Expand")	;,"Expand"
			Load_TvChild(v,TV_FatherID)
			TV_FatherID:=TV_GetParent(TV_FatherID)
		}else{
			;TV_Add(K ":" V,TV_FatherID,"Expand")
		}
	}
}

GetIconCount(file){	;判断文件时候含有图标
	Menu, test, add, test, handle
	Loop
	{
		try {
			id++
			Menu, test, Icon, test, % file, % id
		} catch error {
			break
		}
	}
return id-1
}
handle:
return

AddTab(IconNumber, TabName){
	VarSetCapacity(TCITEM, 100, 0)
	InsertInteger(3, TCITEM, 0)
	InsertInteger(&TabName, TCITEM, 12)
	InsertInteger(IconNumber - 1, TCITEM, 20)
	SendMessage, 0x1307, 999, &TCITEM, SysTabControl321
}

InsertInteger(pInteger, ByRef pDest, pOffset = 0, pSize = 4){
	Loop %pSize%
		DllCall("RtlFillMemory", "UInt", &pDest + pOffset + A_Index-1, "UInt", 1, "UChar", pInteger >> 8*(A_Index-1) & 0xFF)
}
GuiClose:
IL_Destroy(MyImageList)
ExitApp
