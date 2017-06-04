using System;
using System.Data;
using System.Drawing;
using System.Reflection; 
using System.Collections;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.InteropServices; 

using AgentObjects;
using AgentServerObjects;

namespace MicrosoftAgentApplication
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Announcer : System.Windows.Forms.Form
	{
		public const int FAILURE = -1;
		public const int SUCCESS = 100;
		
		private string strPath= AppDomain.CurrentDomain.BaseDirectory ;
		private string strFileName = string.Empty;
		IAgentCharacterEx CharacterEx=null;

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button BtnSpeak;
		private System.Windows.Forms.ComboBox CBSelectStyle;
		private System.Windows.Forms.Label LblEnterText;
		private System.Windows.Forms.TextBox TxtSpeakInput;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Announcer()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used and hide the current agent.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			int dwReqID=0;
			switch(strFileName.ToUpper() )
			{
				case "\\GENIE.ACS": 
					CharacterEx.Hide(0, out dwReqID);
					break;
				case "\\MERLIN.ACS": 
					CharacterEx.Hide(0, out dwReqID);
					break;
				case "\\PEEDY.ACS": 
					CharacterEx.Hide(0, out dwReqID);
					break;
				case "\\ROBBY.ACS": 
					CharacterEx.Hide(0, out dwReqID);
					break;
			}
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.BtnSpeak = new System.Windows.Forms.Button();
			this.CBSelectStyle = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.LblEnterText = new System.Windows.Forms.Label();
			this.TxtSpeakInput = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// BtnSpeak
			// 
			this.BtnSpeak.Location = new System.Drawing.Point(96, 128);
			this.BtnSpeak.Name = "BtnSpeak";
			this.BtnSpeak.TabIndex = 0;
			this.BtnSpeak.Text = "Speak";
			this.BtnSpeak.Click += new System.EventHandler(this.BtnSpeak_Click);
			// 
			// CBSelectStyle
			// 
			this.CBSelectStyle.Items.AddRange(new object[] {
															   "Genie",
															   "Merlin",
															   "Peedy",
															   "Robby"});
			this.CBSelectStyle.Location = new System.Drawing.Point(96, 16);
			this.CBSelectStyle.Name = "CBSelectStyle";
			this.CBSelectStyle.Size = new System.Drawing.Size(192, 21);
			this.CBSelectStyle.TabIndex = 1;
			this.CBSelectStyle.Text = "Select ";
			this.CBSelectStyle.SelectedIndexChanged += new System.EventHandler(this.CBSelectStyle_SelectedIndexChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 24);
			this.label1.TabIndex = 2;
			this.label1.Text = "Select Style";
			// 
			// LblEnterText
			// 
			this.LblEnterText.Location = new System.Drawing.Point(8, 56);
			this.LblEnterText.Name = "LblEnterText";
			this.LblEnterText.Size = new System.Drawing.Size(72, 24);
			this.LblEnterText.TabIndex = 3;
			this.LblEnterText.Text = "Enter Text";
			// 
			// TxtSpeakInput
			// 
			this.TxtSpeakInput.Location = new System.Drawing.Point(96, 48);
			this.TxtSpeakInput.Multiline = true;
			this.TxtSpeakInput.Name = "TxtSpeakInput";
			this.TxtSpeakInput.Size = new System.Drawing.Size(192, 64);
			this.TxtSpeakInput.TabIndex = 4;
			this.TxtSpeakInput.Text = "";
			// 
			// Announcer
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(312, 165);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.TxtSpeakInput,
																		  this.LblEnterText,
																		  this.label1,
																		  this.CBSelectStyle,
																		  this.BtnSpeak});
			this.Name = "Announcer";
			this.Text = "Announcer";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Announcer());
		}

		/// <summary>
		/// Ask agent to speak 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void BtnSpeak_Click(object sender, System.EventArgs e)
		{
			AgentServer Srv = new AgentServer();
			if (Srv == null) 
			{
				MessageBox.Show("ERROR: Agent Server couldn't be started!");
				
			}
			
			IAgentEx SrvEx;
			// The following cast does the QueryInterface to fetch IAgentEx interface from the IAgent interface, directly supported by the object
			SrvEx = (IAgentEx) Srv;
		
			// First try to load the default character
			int dwCharID=0, dwReqID=0;
			try 
			{
				// null is used where VT_EMPTY variant is expected by the COM object
				String strAgentCharacterFile = null;
				if (!strFileName.Equals(string.Empty))  
					//Get the acs path
					strAgentCharacterFile = strPath + strFileName;
				else
				{
					MessageBox.Show("Select Style");
					return;
				}
				
				if (TxtSpeakInput.Text.Equals(string.Empty)) 
				{
					TxtSpeakInput.Text = "Please enter text in textbox";
				}
				SrvEx.Load(strAgentCharacterFile, out dwCharID, out dwReqID);
			} 
			catch (Exception) 
			{
				MessageBox.Show("Failed to load Agent character! Exception details:");
			}
		
			SrvEx.GetCharacterEx(dwCharID, out CharacterEx);
		
			//CharacterEx.SetLanguageID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US));
		
			// Show the character.  The first parameter tells Microsoft
			// Agent to show the character by playing an animation.
			
			// Make the character speak
			// Second parameter will be transferred to the COM object as NULL
				CharacterEx.Speak(TxtSpeakInput.Text, null, out dwReqID);
			
		}

		/// <summary>
		/// Hide current agent Select new agent and show that agent
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void CBSelectStyle_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			
			if (CBSelectStyle.SelectedIndex > -1)
			{	
				HideAgent();
				strFileName =  "\\"+ CBSelectStyle.SelectedItem.ToString() + ".acs";  
				ShowAgent();
			}
 
		}

		/// <summary>
		/// Hide Agent
		/// </summary>
		private void HideAgent()
		{
			int dwReqID=0;
			
			switch(strFileName.ToUpper() )
			{
				case "\\GENIE.ACS": 
					CharacterEx.Hide(0, out dwReqID);
					break;
				case "\\MERLIN.ACS": 
					CharacterEx.Hide(0, out dwReqID);
					break;
				case "\\PEEDY.ACS": 
					CharacterEx.Hide(0, out dwReqID);
					break;
				case "\\ROBBY.ACS": 
					CharacterEx.Hide(0, out dwReqID);
					break;
				default :
					break;
			}
		}

		/// <summary>
		/// Show Agent
		/// </summary>
		private void ShowAgent()
		{
			AgentServer Srv = new AgentServer();
			if (Srv == null) 
			{
				MessageBox.Show("ERROR: Agent Server couldn't be started!");
				
			}
			
			IAgentEx SrvEx;
			// The following cast does the QueryInterface to fetch IAgentEx interface from the IAgent interface, directly supported by the object
			SrvEx = (IAgentEx) Srv;
		
			// First try to load the default character
			int dwCharID=0, dwReqID=0;
			try 
			{
				// null is used where VT_EMPTY variant is expected by the COM object
				String strAgentCharacterFile = null;
				if (!strFileName.Equals(string.Empty))  
					//Get the acs path
					strAgentCharacterFile = strPath + strFileName;
				else
				{
					MessageBox.Show("Select Style");
					return;
				}
				
				if (TxtSpeakInput.Text.Equals(string.Empty)) 
				{
					TxtSpeakInput.Text = "Please enter text in textbox";
				}
				
				//load the acs file
				SrvEx.Load(strAgentCharacterFile, out dwCharID, out dwReqID);
				
			} 
			catch (Exception) 
			{
				MessageBox.Show("Failed to load Agent character! Exception details:");
			}
		
			SrvEx.GetCharacterEx(dwCharID, out CharacterEx);
			//show the agent
			CharacterEx.Show(0, out dwReqID);
			
		}
	}
}
