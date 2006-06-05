function CopyToClipboard(mytxtAreaID)
{
	var mytxtArea = document.getElementById(mytxtAreaID);
	mytxtArea.focus();
  mytxtArea.select();
  if(document.selection.createRange)
  {
		CopiedTxt = document.selection.createRange();
		CopiedTxt.execCommand("Copy");
	}
	else
	{
		//Firefox
	}
}

function PasteFromClipboard(mytxtAreaID)
{ 
	var mytxtArea = document.getElementById(mytxtAreaID);
  mytxtArea.focus();
  mytxtArea.select();
  if(mytxtArea.createTextRange)
  {
		PastedText = mytxtArea.createTextRange();
		PastedText.execCommand("Paste");
	}
	else
	{
		//??? FireFox
	}
} 