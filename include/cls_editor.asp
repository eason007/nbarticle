<!--#Include File="../editor/fck_editor/fckeditor.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：cls_editor.asp
'= 摘    要：编辑器类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-03-29
'====================================================================

Class cls_Editor
	Public iEditorType
	Public iWindowWidth,iWindowHeight
	Public sValue
	
	Private Sub Class_Initialize()
		iEditorType=0
		iWindowWidth="99%"
		iWindowHeight="420"
		sValue=""
	End Sub
	
	Public Property Let EditorType( TypeValue )
		iEditorType = TypeValue
	End Property
	
	Public Property Let Width( widthValue )
		iWindowWidth = widthValue
	End Property

	Public Property Let Height( heightValue )
		iWindowHeight = heightValue
	End Property

	Public Property Let Value( newValue )
		sValue = newValue
	End Property
	
	Public Function Create()
		Dim sOutStr
		
		Select Case iEditorType
		Case 0
		'eWeb Editor
			sOutStr="<input type=hidden name=d_originalfilename>"&VBCrlf
			sOutStr=sOutStr&"<input type=hidden name=d_savefilename>"&VBCrlf
			sOutStr=sOutStr&"<input type=hidden name=d_savepathfilename onchange=""doChange(this,document.a1.d_picture)"">"&VBCrlf
			sOutStr=sOutStr&"<textarea name=""content"" style=""display:none"">"&Server.HTMLEncode(sValue)&"</textarea>"&VBCrlf
			sOutStr=sOutStr&"<iframe ID=""content1"" src="""&SystemFolder&"editor/eweb_editor/ewebeditor.asp?id=content&style=standard&originalfilename=d_originalfilename&savefilename=d_savefilename &savepathfilename=d_savepathfilename"" frameborder=""0"" scrolling=""no"" width="""&iWindowWidth&""" HEIGHT="""&iWindowHeight&"""></iframe>"&VBCrlf
		Case 1
		'FCK Editor
			Dim sBasePath
			sBasePath = SystemFolder&"editor/fck_editor/"

			Dim oFCKeditor
			Set oFCKeditor = New FCKeditor
			
			oFCKeditor.BasePath = sBasePath
			oFCKeditor.Width=iWindowWidth
			oFCKeditor.Height=iWindowHeight
			oFCKeditor.Value=sValue
			oFCKeditor.Config("AutoDetectLanguage") = False
			oFCKeditor.Config("DefaultLanguage")    = "zh-cn"
			oFCKeditor.Value = sValue
			
			sOutStr=oFCKeditor.Create("content")
		Case 2
		'InnovaStudio Editor
			sOutStr = "<script language=JavaScript src='"&SystemFolder&"editor/innova_editor/scripts/language/schi/editor_lang.js'></script>"&VBCrlf
			If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"MSIE") Then
				sOutStr = sOutStr & "<script language=JavaScript src='"&SystemFolder&"editor/innova_editor/scripts/editor.js'></script>"&VBCrlf
			Else
				sOutStr = sOutStr & "<script language=JavaScript src='"&SystemFolder&"editor/innova_editor/scripts/moz/editor.js'></script>"&VBCrlf
			End If
			sOutStr = sOutStr & "<pre id=""idTemporary"" name=""idTemporary"" style=""display:none"">"&Server.HTMLEncode(sValue)&"</pre>"&VBCrlf
			sOutStr = sOutStr & "<script>"&VBCrlf
			sOutStr = sOutStr & "var oEdit1 = new InnovaEditor(""oEdit1"");"&VBCrlf
			sOutStr = sOutStr & "oEdit1.cmdAssetManager=""modalDialogShow('"&SystemFolder&"editor/innova_editor/assetmanager/assetmanager.asp?lang=schi',640,465)"";"&VBCrlf
			sOutStr = sOutStr & "oEdit1.btnFlash=true;"&VBCrlf
			sOutStr = sOutStr & "oEdit1.btnMedia=true;"&VBCrlf

			sOutStr = sOutStr & "oEdit1.RENDER(document.getElementById(""idTemporary"").innerHTML);"&VBCrlf
			sOutStr = sOutStr & "</script>"&VBCrlf
			sOutStr = sOutStr & "<input type=""hidden"" name=""content"" id=""content"">"&VBCrlf
		Case 3
		'TinyMCE Editor
			sOutStr = "<script language=""javascript"" type=""text/javascript"" src=""" & SystemFolder & "editor/tinymce_editor/tiny_mce.js""></script>"&VBCrlf
			sOutStr = sOutStr & "<script language=""javascript"" type=""text/javascript"">"&VBCrlf
			sOutStr = sOutStr & "tinyMCE.init({"&VBCrlf
			sOutStr = sOutStr & "	mode : ""exact"","&VBCrlf
			sOutStr = sOutStr & "	theme : ""advanced"","&VBCrlf
			sOutStr = sOutStr & "	elements : ""content"","&VBCrlf
			sOutStr = sOutStr & "	width : """&iWindowWidth&""","&VBCrlf
			sOutStr = sOutStr & "	height : """&iWindowHeight&""","&VBCrlf
			sOutStr = sOutStr & "	plugins : ""table,save,advhr,advimage,advlink,emotions,iespell,insertdatetime,preview,zoom,flash,searchreplace,print,contextmenu,paste,directionality,fullscreen"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_buttons1_add_before : ""save,newdocument,separator"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_buttons1_add : ""fontselect,fontsizeselect"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_buttons2_add : ""separator,insertdate,inserttime,preview,zoom,separator,forecolor,backcolor"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_buttons2_add_before: ""cut,copy,paste,pastetext,pasteword,separator,search,replace,separator"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_buttons3_add_before : ""tablecontrols,separator"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_buttons3_add : ""emotions,iespell,flash,advhr,separator,print,separator,ltr,rtl,separator,fullscreen"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_toolbar_location : ""top"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_toolbar_align : ""left"","&VBCrlf
			sOutStr = sOutStr & "	content_css : ""example_word.css"","&VBCrlf
			sOutStr = sOutStr & "	plugin_insertdate_dateFormat : ""%Y-%m-%d"","&VBCrlf
			sOutStr = sOutStr & "	plugin_insertdate_timeFormat : ""%H:%M:%S"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_resizing : true,"&VBCrlf
			sOutStr = sOutStr & "	extended_valid_elements : ""a[name|href|target|title|onclick],img[class|src|border=0|alt|title|hspace|vspace|width|height|align|onmouseover|onmouseout|name],hr[class|width|size|noshade],font[face|size|color|style],span[class|align|style]"","&VBCrlf
			sOutStr = sOutStr & "	theme_advanced_resizing : true"&VBCrlf
			sOutStr = sOutStr & "});"&VBCrlf
			sOutStr = sOutStr & "</script>"
			sOutStr = sOutStr & "<textarea name=""content"">"&Server.HTMLEncode(sValue)&"</textarea>"
		End Select
		
		Create = sOutStr
	End Function
	
End Class
%>