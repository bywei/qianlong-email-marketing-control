function DrawImage(ImgD,width_s,height_s){
var image=new Image();
image.src=ImgD.src;
if(image.width>0 && image.height>0){
flag=true;
if(image.width/image.height>=width_s/height_s){
if(image.width>width_s){
ImgD.width=width_s;
ImgD.height=(image.height*width_s)/image.width;
}else{
ImgD.width=image.width;
ImgD.height=image.height;
}
}
else{
if(image.height>height_s){
ImgD.height=height_s;
ImgD.width=(image.width*height_s)/image.height;
}else{
ImgD.width=image.width;
ImgD.height=image.height;
}
}
}
}


function Checked()
{
	var j = 0
	for(i=0;i < document.form.elements.length;i++){
		if(document.form.elements[i].name == "FileId" || document.form.elements[i].name == "FolderId"){
			if(document.form.elements[i].checked){
				j++;
			}
		}
	}
	return j;
}
function CheckAll1()
{
	for(i=0;i<document.form.elements.length;i++)
	{
		if(document.form.elements[i].checked){
			document.form.elements[i].checked=false;
			document.form.CheckAll.checked=false;
		}
		else{
			document.form.elements[i].checked = true;
			document.form.CheckAll.checked = true;
		}
	}
}
function DelAll()
{
	if(Checked()  <= 0){
		alert("������ѡ�����е�һ���ļ�");
	}	
	else{
		if(confirm("ȷ��Ҫɾ��ѡ����ļ�ô��\n�˲��������Իָ���")){
			form.action="?action=Del";
			form.submit();
		}
	}
}
function CopyLink(url) {
var l = url
window.clipboardData.setData("text",l);
alert("�����ַ�ѿ����������壬����ʹ��ճ����Ctrl+V���á�\n\n"+l+"")
}

function CopyAll()
{
	if(Checked()  <= 0){
		alert("������ѡ�����е�һ���ļ�");
	}	
	else{
		form.action="?action=CPA";
		form.submit();
	}
}

function MoveAll()
{
	if(Checked()  <= 0){
		alert("������ѡ�����е�һ���ļ�");
	}	
	else{
		if(confirm("ȷ��Ҫ�ƶ�ѡ����ļ�ô��")){
			form.action="?action=YDWJ";
			form.submit();
		}
	}
}


function Rname()
{
	if(Checked() == 0){
		alert("������ѡ��һ���ļ�");
	}
	else{
		if(Checked() != 1){
			alert("ֻ��ѡ��һ���ļ���һ���ļ���");
		}
		else{
			for(i=0;i < document.form.elements.length;i++){
				if(document.form.elements[i].name == "FolderId" && document.form.elements[i].checked){
					var j = prompt("���������ļ�����",document.form.elements[i].value)
					break;
				}
				else if(document.form.elements[i].name == "FileId" && document.form.elements[i].checked){
					var j = prompt("���������ļ���",document.form.elements[i].value)
					break;
				}
			}
			if(j != "" && j != null){
				if(IsStr(j) == j.length){
					form.action="?action=Rname&NewName=" + j;
                                                      form.target="_self";
					form.submit();
				}
				else{
					alert("�����Ʋ����ϱ�׼��ֻ������ĸ�����֡�����»��ߵ����,\n���ܺ��к��֡��ո����������");
				}
			}
		}
	}
}
function IsStr(w)
{
	var str = "abcdefghijklmnopqrstuvwxyz_1234567890."
	 w = w.toLowerCase();
	var j = 0;
	for(i=0;i < w.length;i++){
		if(str.indexOf(w.substr(i,1)) != -1){
			j++;
		}
	}
	return j;
}
function NewFile(form,i)
{
	if(i == 1){
		if(form.NewFolderName.value == ""){
			alert("�ļ���������Ϊ��");
		}
		else{
			if(IsStr(form.NewFolderName.value) == form.NewFolderName.value.length){
				form.action="?action=NewFolder";
				form.submit();
			}
			else{
				alert("�ļ����������ϱ�׼��ֻ������ĸ�����֡�����»��ߵ����,\n���ܺ��к��֡��ո����������");
			}
		}
	}
	else{
		if(form.NewFileName.value == ""){
			alert("�ļ�������Ϊ��");
		}
		else{
			if(IsStr(form.NewFileName.value) == form.NewFileName.value.length){
				form.action="?action=NewFile";
				form.submit();
			}
			else{
				alert("�ļ��������ϱ�׼��ֻ������ĸ�����֡�����»��ߵ����,\n���ܺ��к��֡��ո����������");
			}
		}
	}
}
