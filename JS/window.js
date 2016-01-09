function goWindow(strURL,strName,strWidth,strHeight,strOthers)
{
	var strFeatures = "width=" + strWidth + ",height=" + strHeight;
	if (strOthers.length != 0)
		strFeatures += "," + strOthers;
			
	var newWindow = window.open(strURL,strName,strFeatures);
	newWindow.focus();
}  