function ChangeCodeEditProduct(ctx)
{
	if (dnn.dom.getById(ctx + '_chkDirectCodeEdit').checked)
	{
		dnn.dom.getById(ctx + '_txtProductVCode').readOnly = false;
		dnn.dom.getById(ctx + '_txtProductGCode').readOnly = false;
		dnn.dom.getById(ctx + '_txtProductVCode').style.backgroundColor="#FFFFFF";
		dnn.dom.getById(ctx + '_txtProductGCode').style.backgroundColor="#FFFFFF";
		dnn.dom.getById(ctx + '_txtProductVCode').style.color="#000000";
		dnn.dom.getById(ctx + '_txtProductGCode').style.color="#000000";
	}
	else
	{
		dnn.dom.getById(ctx + '_txtProductVCode').readOnly = true;
		dnn.dom.getById(ctx + '_txtProductGCode').readOnly = true;
		dnn.dom.getById(ctx + '_txtProductVCode').style.backgroundColor="#DBDBDB";
		dnn.dom.getById(ctx + '_txtProductGCode').style.backgroundColor="#DBDBDB";
		dnn.dom.getById(ctx + '_txtProductVCode').style.color="#424242";
		dnn.dom.getById(ctx + '_txtProductGCode').style.color="#424242";
	}
}
function ChangeCodeEditCode(ctx)
{
	if (dnn.dom.getById(ctx + '_chkDirectCodeEdit').checked)
	{
		dnn.dom.getById(ctx + '_txtActivationCode').readOnly = false;
		dnn.dom.getById(ctx + '_txtUserName').readOnly = false;
		dnn.dom.getById(ctx + '_txtActivationCode').style.backgroundColor="#FFFFFF";
		dnn.dom.getById(ctx + '_txtUserName').style.backgroundColor="#FFFFFF";
		dnn.dom.getById(ctx + '_txtActivationCode').style.color="#000000";
		dnn.dom.getById(ctx + '_txtUserName').style.color="#000000";
	}
	else
	{
		dnn.dom.getById(ctx + '_txtActivationCode').readOnly = true;
		dnn.dom.getById(ctx + '_txtUserName').readOnly = true;
		dnn.dom.getById(ctx + '_txtActivationCode').style.backgroundColor="#DBDBDB";
		dnn.dom.getById(ctx + '_txtUserName').style.backgroundColor="#DBDBDB";
		dnn.dom.getById(ctx + '_txtActivationCode').style.color="#424242";
		dnn.dom.getById(ctx + '_txtUserName').style.color="#424242";
	}
}

function successFuncEmail(result, ctx)
{
	DoRedirect(result);
}

function successFuncGenerate(result, ctx)
{
	var string_array=result.split("|");

	dnn.dom.getById(ctx + '_txtProductVCode').value = string_array[2];
	dnn.dom.getById(ctx + '_txtProductGCode').value = string_array[3];

	EnableAllControlsProduct(ctx);
	
	dnn.dom.getById(ctx + '_btnGenerateNew').style.cursor = 'default';
	document.body.style.cursor = 'default';
	
	enableElement(ctx + '_btnClearCodes');
	createPopup(ctx, 'Generation', "Product code generation succesfull!<br><br><img src='" + modulePath + "button_ok.gif' style='cursor:pointer;cursor:hand;' onclick='hidebox();return false'>", 250, 200);
}

function successFuncValidate(result, ctx)
{
	if (result=='true') 
	{
		createPopup(ctx, 'Validation', "Validation succesfull!", 250, 200);
		//alert("Validation succesfull!");
	}
	else
	{
		createPopup(ctx, 'Validation', "Validation failure! Plese correct product codes.", 250, 200);
		//alert("Validation failure! Plese correct product codes.");
	}

	EnableAllControlsProduct(ctx);

	dnn.dom.getById(ctx + '_btnValidate').style.cursor = 'default';
	document.body.style.cursor = 'default';
}

function successFuncActivateClient(result, ctx)
{
	dnn.dom.getById(ctx + '_txtActivationCode').value = result;
	dnn.dom.getById(ctx + '_btnGenerate').style.cursor = 'default';
	document.body.style.cursor = 'default';
	EnableAllControlsCodeClient(ctx);
	if (result!='')
	{
		hideElement(ctx + '_btnGenerate');
		createPopup(ctx, 'Activation', "License key was generated succesfully.<br><br><img src='" + modulePath + "button_ok.gif' style='cursor:pointer;cursor:hand;' onclick='hidebox();return false'>", 250, 200);
	}
	else
	{
		createPopup(ctx, 'Activation', "Error in code generation.<br><br><img src='" + modulePath + "button_ok.gif' style='cursor:pointer;cursor:hand;' onclick='hidebox();return false'>", 250, 200);
	}
}

function successFuncActivate(result, ctx)
{
	dnn.dom.getById(ctx + '_txtActivationCode').value = result;
	dnn.dom.getById(ctx + '_btnGenerate').style.cursor = 'default';
	document.body.style.cursor = 'default';
	EnableAllControlsCode(ctx);
	if (result!=null)
	{
		createPopup(ctx, 'Activation', "License key was generated succesfully.<br><br><img src='" + modulePath + "button_ok.gif' style='cursor:pointer;cursor:hand;' onclick='hidebox();return false'>", 250, 200);
	}
	else
	{
		createPopup(ctx, 'Activation', "Error in code generation.<br><br><img src='" + modulePath + "button_ok.gif' style='cursor:pointer;cursor:hand;' onclick='hidebox();return false'>", 250, 200);
	}
}

function successFuncGetUser(result, ctx)
{
	dnn.dom.getById(ctx + '_txtUserName').value = result;
}

function keepvalue(mycontrol, myvalue)
{
	handled=true
	dnn.dom.getById(mycontrol).value = myvalue;
}

function errorFunc(result, ctx)
{
	createPopup(ctx, 'Error', result, 250, 200);
	//EnableAllControlsProduct(ctx);
	document.body.style.cursor = 'default';
	//alert('Error: ' + result);
}

function showHide(id)
{
el = document.getElementById(id); 
     if (el.style.display == 'none')
        { 
            el.style.display = 'block'; 
        } 
     else 
        { 
            el.style.display = 'none'; 
        } 
} 

function showElement(id)
{
	el = document.getElementById(id); 
	el.style.display = 'block'; 
	//el.style.visibility = 'show';
} 

function hideElement(id)
{
	el = document.getElementById(id); 
	el.style.display = 'none'; 
	//el.style.visibility = 'hidden';
} 

function enableElement(id)
{
el = document.getElementById(id); 
el.disabled = false;
} 

function disableElement(id)
{
el = document.getElementById(id); 
el.disabled = true;
} 

function pausecomp(millis) 
{
	date = new Date();
	var curDate = null;
	
	do { var curDate = new Date(); } 
	while(curDate-date < millis);
} 

function CopyToClipboard(mytxtAreaID)
{
	var mytxtArea = dnn.dom.getById(mytxtAreaID);
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
	var mytxtArea = dnn.dom.getById(mytxtAreaID);
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

function ClearCodes(ctx)
{
	dnn.dom.getById(ctx + '_txtProductVCode').value = '';
	dnn.dom.getById(ctx + '_txtProductGCode').value = '';
	enableElement(ctx + '_btnGenerateNew');
	disableElement(ctx + '_btnClearCodes');
}

function EnableAllControlsProduct(ctx)
{
	enableElement(ctx + '_btnGenerateNew');
	enableElement(ctx + '_btnClearCodes');
	enableElement(ctx + '_btnValidate');
}

function DisableAllControlsProduct(ctx)
{
	disableElement(ctx + '_btnGenerateNew');
	disableElement(ctx + '_btnClearCodes');
	disableElement(ctx + '_btnValidate');
}
function DisableAllControlsCode(ctx)
{
	disableElement(ctx + '_btnGenerate');
	disableElement(ctx + '_imgEmail');
}
function DisableAllControlsCodeClient(ctx)
{
	disableElement(ctx + '_btnGenerate');
}
function EnableAllControlsCode(ctx)
{
	enableElement(ctx + '_btnGenerate');
	enableElement(ctx + '_imgEmail');
}
function EnableAllControlsCodeClient(ctx)
{
	enableElement(ctx + '_btnGenerate');
}

function ShowProducts(ctx)
{
	dnn.dom.getById(ctx + '_txtLastPage').value = 'p';
	showElement(ctx + '_pnlProducts');
	hideElement(ctx + '_pnlCustomers');
	hideElement(ctx + '_pnlCodes');
	//dnn.dom.getById(ctx + '_tdbtnProducts').style.border.bottom = "solid 0px orange";
	dnn.dom.getById(ctx + '_tdbtnProducts').className = "multipageactive";
	dnn.dom.getById(ctx + '_tdbtnCustomers').className = "multipage";
	dnn.dom.getById(ctx + '_tdbtnCodes').className = "multipage";
}

function ShowCustomers(ctx)
{
	dnn.dom.getById(ctx + '_txtLastPage').value = 'c';
	hideElement(ctx + '_pnlProducts');
	showElement(ctx + '_pnlCustomers');
	hideElement(ctx + '_pnlCodes');
	
	dnn.dom.getById(ctx + '_tdbtnProducts').className = "multipage";
	dnn.dom.getById(ctx + '_tdbtnCustomers').className = "multipageactive";
	dnn.dom.getById(ctx + '_tdbtnCodes').className = "multipage";
}

function ShowCodes(ctx)
{
	dnn.dom.getById(ctx + '_txtLastPage').value = 'd';
	hideElement(ctx + '_pnlProducts');
	hideElement(ctx + '_pnlCustomers');
	showElement(ctx + '_pnlCodes');
	dnn.dom.getById(ctx + '_tdbtnProducts').className = "multipage"
	dnn.dom.getById(ctx + '_tdbtnCustomers').className = "multipage"
	dnn.dom.getById(ctx + '_tdbtnCodes').className = "multipageactive"
}

function ShowLicenseType(ctx, ctxform)
{
	var mVal = dnn.dom.getById(ctx).value;
	switch (mVal)
	{
		case "0":
			showElement(ctxform + '_txtExpireDate');
			showElement(ctxform + '_imgExpireDateCalendar');
			//showElement(ctxform + '_plExpireDate_label');
			showElement(ctxform + '_pnlExpireDate');
			//hideElement(ctxform + '_plPeriodicDays_label');
			hideElement(ctxform + '_pnlPeriodicDays');
			break;
		case "1":
			showElement(ctxform + '_txtExpireDate');
			hideElement(ctxform + '_imgExpireDateCalendar');
			//hideElement(ctxform + '_plExpireDate_label');
			hideElement(ctxform + '_pnlExpireDate');
			//showElement(ctxform + '_plPeriodicDays_label');
			showElement(ctxform + '_pnlPeriodicDays');
			break;
		case "2":
			hideElement(ctxform + '_txtExpireDate');
			hideElement(ctxform + '_imgExpireDateCalendar');
			//hideElement(ctxform + '_plExpireDate_label');
			hideElement(ctxform + '_pnlExpireDate');
			//hideElement(ctxform + '_plPeriodicDays_label');
			hideElement(ctxform + '_pnlPeriodicDays');
			break;
		default:
			break;
	}
}

function GenByCustomer(ctx, ctxform)
{
	if (!dnn.dom.getById(ctx).checked)
	{
		showElement(ctxform + '_pnlGenByCustomer');
	}
	else
	{
		hideElement(ctxform + '_pnlGenByCustomer');
	}
}

function DoRedirect(msg)
{ 
location.href = msg;
return true; 
} 
