using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Web.UI.Design;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;


namespace msWebControlsLibrary
{
	/// <summary>
	/// Summary description for foImageButton.
	/// </summary>
	[
	ToolboxData("<{0}:ExImageButton runat=server></{0}:ExImageButton>"),
	Designer(typeof(ExImageButtonDesigner))
	]
	public class ExImageButton : System.Web.UI.WebControls.WebControl, INamingContainer
	{
		private string _strDisableImageURL;
		private string _strMouseOverImageURL;
		private string _strMouseOutImageURL;
		private System.Web.UI.WebControls.ImageButton btnImage;
		private System.Web.UI.WebControls.Image btnDisableImage;
		// The sole event hook that the button exposes.
		public event EventHandler Click;
		#region DesignProperties
		[
		Browsable(true),
		Category("Behavior") ] 
		public bool CausesValidation
		{
			get 
			{
			
				return btnImage.CausesValidation;  
			}
			set
			{
			
				btnImage.CausesValidation=value;
			}
		}
		
		[
		Bindable(true), 
		Browsable(true),
		Category("Appearance") ] 
		public string DisableImageURL 
		{
			get
			{
				return _strDisableImageURL;
			}

			set
			{
				_strDisableImageURL = value;
				DisableImageURLFromViewState=_strDisableImageURL;
			}
		}

		[
		Bindable(false), 
		Browsable(true),
		Category("Appearance") ] 
		public string MouseOverImageURL 
		{
			get
			{
				return _strMouseOverImageURL;
			}

			set
			{
				_strMouseOverImageURL = value;
				MouseOverImageURLFromViewState=_strMouseOverImageURL;
			}
		}

		[
		Bindable(false), 
		Browsable(true),
		Category("Appearance") ] 
		public string MouseOutImageURL 
		{
			get
			{
				return _strMouseOutImageURL;
			}

			set
			{
				_strMouseOutImageURL = value;
				MouseOutImageURLFromViewState=_strMouseOutImageURL;
			}
		}

		[Bindable(false), Category("Appearance"), DefaultValue(null), 
		Description("The text that is displayed inside the button.")] 
		public string Text 
		{
			get 
			{
				this.EnsureChildControls();

				return btnImage.AlternateText;
			}

			set 
			{
				this.EnsureChildControls();
				btnImage.AlternateText = value;
			}
		}

		[Bindable(false), Category("Appearance"), DefaultValue(""),
		Description("The URL of the image to be displayed for the button.")] 
		public string ImageUrl 
		{
			get 
			{
				this.EnsureChildControls();
				return btnImage.ImageUrl;
			}

			set 
			{
				this.EnsureChildControls();
				btnImage.ImageUrl = value;
			}
		}
		#endregion
		#region DesignPropertiesViewStates
		protected string DisableImageURLFromViewState
		{
			get
			{
				return (string) ViewState["DisableImageURLFromViewState"];
			}
			set
			{
				ViewState["DisableImageURLFromViewState"] = value;
			}
		
		}
		protected string MouseOverImageURLFromViewState
		{
			get
			{
				return (string) ViewState["MouseOverImageURLFromViewState"];
			}
			set
			{
				ViewState["MouseOverImageURLFromViewState"] = value;
			}
		
		}
		protected string MouseOutImageURLFromViewState
		{
			get
			{
				return (string) ViewState["MouseOutImageURLFromViewState"];
			}
			set
			{
				ViewState["MouseOutImageURLFromViewState"] = value;
			}
		
		}
		#endregion
		
		override protected void OnInit(EventArgs ec) 
		{
			base.OnInit(ec);
		}

		override protected void OnLoad(EventArgs eo) 
		{
			
		}

		override protected void OnPreRender(EventArgs ecx) 
		{
			try 
			{
				btnDisableImage.ImageUrl=DisableImageURLFromViewState;
				btnImage.Attributes.Remove("OnMouseOver");
				btnImage.Attributes.Remove("OnMouseOut");
			} 
			finally 
			{
				if(MouseOverImageURLFromViewState!=null) btnImage.Attributes.Add("OnMouseOver", "this.src='"+ MouseOverImageURLFromViewState  +"';");
				
				if(MouseOutImageURLFromViewState!=null) btnImage.Attributes.Add("OnMouseOut", "this.src='"+ MouseOutImageURLFromViewState +"';");
				
			}
			if(Enabled)
			{
				btnImage.Visible=true;
				btnDisableImage.Visible=false; 
			}
			else
			{
				btnImage.Visible=false;
				btnDisableImage.Visible=true; 
			}
			
		}

		protected override void CreateChildControls() 
		{
			// Setup the controls on the page.
			btnImage = new System.Web.UI.WebControls.ImageButton();
			 
			btnDisableImage =new System.Web.UI.WebControls.Image();

			btnDisableImage.ImageUrl=DisableImageURLFromViewState;
			btnDisableImage.Visible=false; 
 
			// Setup the events on the page.
			btnImage.Click += new ImageClickEventHandler(this.OnBtnImage_Click);
			Controls.Add(btnImage);
			Controls.Add(btnDisableImage);

			
			
		}
		
		protected virtual void OnBtnImage_Click(object sender, ImageClickEventArgs e)
		{
			if (Click != null && Enabled ) 
			{
				Click(this, e);
			}  
		}
		

	}

	public class ExImageButtonDesigner : ControlDesigner
	{
		public override string GetDesignTimeHtml()
		{
			ExImageButton cb = (ExImageButton) Component;

			if(cb.Controls.Count == 0)
				cb.Text = cb.UniqueID;
				
			StringWriter sw = new StringWriter();
			HtmlTextWriter tw = new HtmlTextWriter(sw);

			cb.RenderBeginTag(tw);
			cb.RenderControl(tw);
			cb.RenderEndTag(tw);
					
			return(sw.ToString());	
		}
	}

}