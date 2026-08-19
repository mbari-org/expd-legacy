<%@ Page Language="VB" ContentType="text/html" Trace="False" EnableViewState="true" %>

<%@ Register TagPrefix="DP" Assembly="RichDatePicker" Namespace="RicherComponents.DatePickerControl" %>

<Script Runat="Server">
'*****************************************************************************
'		
'*****************************************************************************	
Sub Page_Load

  'turn off all validators so that all button callbacks can work.
  'we will turn them on in button callbacks that need form validation.
  DisableValidators()
  
  If Page.IsPostBack Then
  
    Trace.Write("<br> postback <br>")
	
	select case ddlRequestBy.SelectedItem.Value
	  case "dive" 'dive
	    pnlDive.Visible = true
        pnlYearday.Visible = false
	    pnlTimeRange.Visible = false
	  case "yearday" 'year/yeaday
	    pnlDive.Visible = false
        pnlYearday.Visible = true
	    pnlTimeRange.Visible = false
	  case "timerange" 'start/end datetime
	    pnlDive.Visible = false
        pnlYearday.Visible = false
	    pnlTimeRange.Visible = true
    end select
	
	'lstVehicle.SelectedIndex = lstVehicle.Items.IndexOf(lstVehicle.Items.FindByValue("vnta"))

  Else
    Trace.Write("<br> post <br>")
	 lstReturn.SelectedIndex = lstReturn.Items.IndexOf(lstReturn.Items.FindByValue("data1"))
	 lstVehicle.SelectedIndex = lstVehicle.Items.IndexOf(lstVehicle.Items.FindByValue("tibr"))
	 lstDatasource.SelectedIndex = lstDatasource.Items.IndexOf(lstDatasource.Items.FindByValue("binned"))
    'tableOne.Visible = false
  End If
	

End Sub 'Page_Load()


'*****************************************************************************
'		
'*****************************************************************************	
Sub DisableValidators()
      dim bv as BaseValidator
      for each bv in Validators
        bv.Enabled = false
      next
end sub

'*****************************************************************************
'		
'*****************************************************************************		
Sub EnableValidators(container as Control)
	'Sub EnableValidators(ByVal container As Object)
	    'dim container as Control
        dim c as Control 
		dim bv as BaseValidator
		Trace.Write("<br> enable validation <br>")
		             
       for each c in container.Controls
           If TypeOf c Is IValidator Then         
          'if(c is IValidator) then
          
            ' Assumption here: 
            ' every control that implements IValidator is derived from BaseValidator class
            '((BaseValidator)c).Enabled = true
			bv = CType(c,BaseValidator)
			bv.Enabled = true
          end if
       next
end sub     

'*****************************************************************************
' I want ending datetime to be greater that starting date time		
'*****************************************************************************	
Sub ServerValidateDateTimes (sender As Object, value As ServerValidateEventArgs)
		    dim theStart as DateTime
			Trace.Write("<br> do validation <br>")
            theStart = dtpStart.SelectedDate
		    dim theEnd as DateTime
            theEnd = dtpEnd.SelectedDate
            Try
                if ( theEnd > theStart ) then
                    value.IsValid = True
                    Exit Sub
                End If
            Catch exc As Exception
            End Try
            value.IsValid = False
   End Sub
  

'*****************************************************************************
' 		
'*****************************************************************************	  
Sub SubmitBtn_Click(Sender As Object, E As EventArgs)
   
     EnableValidators(ValidatorsGroup1)
	   Trace.Write( "SubmitBtn_Click  try validate"  )
	   Page.Validate()
	   
	   
	   
	  
	   if NOT Page.IsValid then
	      Trace.Write( "SubmitBtn_Click  validate failed"  )
	
	   Else
	      Trace.Write( "SubmitBtn_Click  validate OK"  )
	      
	      dim sUrl as string
	      
	    if (lstReturn.SelectedIndex = lstReturn.Items.IndexOf(lstReturn.Items.FindByValue("plot1")) ) then
		     sUrl = "http://mbari1548-COS-vm:8081/rovctd/pages/timeseries.jsp?"
		  else if (lstReturn.SelectedIndex = lstReturn.Items.IndexOf(lstReturn.Items.FindByValue("plot2")) ) then
		     sUrl = "http://mbari1548-COS-vm:8081/rovctd/pages/stackedtimeseries.jsp?"
		  else if (lstReturn.SelectedIndex = lstReturn.Items.IndexOf(lstReturn.Items.FindByValue("plot3")) ) then
		     sUrl = "http://mbari1548-COS-vm:8081/rovctd/pages/profile.jsp?"
		  else if (lstReturn.SelectedIndex = lstReturn.Items.IndexOf(lstReturn.Items.FindByValue("data1")) ) then
		     sUrl = "http://mbari1548-COS-vm:8081/rovctd/servlet/RovctdDataServlet?noHTMLHeader=0&"
		  else  'default
		     sUrl = "http://mbari1548-COS-vm:8081/rovctd/RovDataServlet?noHTMLHeader=0&"
		  end if   
	      
	    
		  
		  if (lstDatasource.SelectedIndex = lstDatasource.Items.IndexOf(lstDatasource.Items.FindByValue("raw")) ) then
		     sUrl=sUrl + "&raw=1"
		  else if (lstDatasource.SelectedIndex = lstDatasource.Items.IndexOf(lstDatasource.Items.FindByValue("binned")) ) then
		     sUrl=sUrl + "&raw=0"
		  else  'default
		     sUrl=sUrl + "&raw=0"
		  end if 
	  
	
		  
		  
		  if (lstVehicle.SelectedIndex = lstVehicle.Items.IndexOf(lstVehicle.Items.FindByValue("tibr")) ) then
		     sUrl=sUrl + "&platform=tibr"
		  else if (lstVehicle.SelectedIndex = lstVehicle.Items.IndexOf(lstVehicle.Items.FindByValue("vnta")) ) then
		     sUrl=sUrl + "&platform=vnta"
		  else if (lstVehicle.SelectedIndex = lstVehicle.Items.IndexOf(lstVehicle.Items.FindByValue("docr")) ) then
		     sUrl=sUrl + "&platform=docr"
		  else  'default
		     sUrl=sUrl + "&platform=vnta"
		  end if 
	
	      
		      
	      select case ddlRequestBy.SelectedItem.Value
	        case "dive" 'dive
			   sUrl=sUrl + "&dive=" + trim(txtDive.text)
	        case "yearday" 'year/yeaday
	           sUrl=sUrl + "&year=" + trim(txtYear.text)
			   sUrl=sUrl + "&yearday=" + trim(txtYearday.text)
	        case "timerange" 'start/end datetime
	        
	        dim theStart as DateTime
          theStart = dtpStart.SelectedDate
		      dim theEnd as DateTime
           theEnd = dtpEnd.SelectedDate
	   
	        ''theStart.toString("dd-MMM-yyyy HH:mm:ss")
	        
			   sUrl=sUrl + "&start=" + Server.urlEncode(theStart.toString("dd-MMM-yyyy HH:mm:ss"))
			   sUrl=sUrl + "&end=" + Server.urlEncode(theEnd.toString("dd-MMM-yyyy HH:mm:ss"))
	        
        end select
		    
		    
		    
		    
			  if (lstReturn.SelectedIndex = lstReturn.Items.IndexOf(lstReturn.Items.FindByValue("plot2")) ) then
			    sUrl=sUrl + "&domain=epochsecs"
			  sUrl=sUrl + "&r1=p"
			  sUrl=sUrl + "&r2=t"
			  sUrl=sUrl + "&r3=s"
			  sUrl=sUrl + "&r4=light"
			  sUrl=sUrl + "&r5=o2"
			  else if (lstReturn.SelectedIndex = lstReturn.Items.IndexOf(lstReturn.Items.FindByValue("plot3")) ) then
			  sUrl=sUrl + "&domain=p"
			  sUrl=sUrl + "&r1=t"
			  sUrl=sUrl + "&r2=s"
			  sUrl=sUrl + "&r3=o2"
			  sUrl=sUrl + "&r4=light"
			  else
			  sUrl=sUrl + "&domain=epochsecs"
			  sUrl=sUrl + "&r1=p"
			  sUrl=sUrl + "&r2=t"
			  sUrl=sUrl + "&r3=s"
			  sUrl=sUrl + "&r4=light"
			  sUrl=sUrl + "&r5=o2"
			  
			  end if
			  
	          Response.Redirect(sUrl)	      
	   End If 'is page valid
		 
	   Trace.Write( "SubmitBtn_Click done"  )
   
    'Response.write(sUrl)

End Sub	

</Script>

<html >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>
<body>

<form runat="server">

  <table width="500" border="0">
    <tr>
      <td width="76">Platform:</td>
      <td width="408">
	  
		<asp:RadioButtonList ID="lstVehicle"
		runat="server" RepeatDirection="Horizontal" >
		<asp:ListItem value="vnta">Ventana</asp:ListItem>
		<asp:ListItem value="tibr">Tiburon</asp:ListItem>
		<asp:ListItem value="docr">DocRicketts</asp:ListItem>
	  </asp:RadioButtonList></td>
    </tr>
  </table>
  
  <table width="500" border="0">
    <tr>
      <td width="76">Datasource:</td>
      <td width="408">
	  
		<asp:RadioButtonList ID="lstDatasource"
		runat="server" RepeatDirection="Horizontal" >
		<asp:ListItem value="raw">Raw</asp:ListItem>
		<asp:ListItem value="binned">15 sec binned</asp:ListItem>
	  </asp:RadioButtonList></td>
    </tr>
  </table>
  
  <table width="500" border="0">
    <tr>
      <td width="76">Return:</td>
      <td width="408">
	  
		<asp:RadioButtonList ID="lstReturn"
		runat="server" RepeatDirection="Horizontal">
		<asp:ListItem value="plot1">Multiline Timeseries Plot</asp:ListItem>
		<asp:ListItem value="plot2">Stacked Timeseries Plot</asp:ListItem>
		<asp:ListItem value="plot3">Profile Plot</asp:ListItem>
		<asp:ListItem value="data1">DataList</asp:ListItem>
	  </asp:RadioButtonList></td>
    </tr>
  </table>
  
  
  
  <p>Request By:
    <asp:DropDownList ID="ddlRequestBy" runat="server" AutoPostBack="true">
	<asp:ListItem value="dive">Dive Number</asp:ListItem>
	<asp:ListItem value="yearday">Year/Yday</asp:ListItem>
	<asp:ListItem value="timerange">Start/End Times</asp:ListItem>
	</asp:DropDownList>
  </p>
  


  <asp:Panel id="pnlDive" runat="server" Visible="True">
    <p>
    <asp:Label ID="lblDive" Text="Dive:" runat="server" />        
    <asp:TextBox ID="txtDive" runat="server" /> 
	</p>
  </asp:Panel> 

  <asp:Panel id="pnlYearday" runat="server" Visible="False">
    <p>
    <asp:Label ID="lblYear" Text="Year:" runat="server" />    
    <asp:TextBox ID="txtYear" runat="server" />
    <asp:Label ID="lblYearday" Text="YearDay:" runat="server" />    
    <asp:TextBox ID="txtYearday" runat="server" />  </p>
 	</p>
  </asp:Panel> 
  

   <asp:Panel id="pnlTimeRange" runat="server" Visible="False">
   <asp:Placeholder id="ValidatorsGroup1" runat="server">
   <table width="509" border="0" >
   <tr>
    <td width="118" align="left" valign="top" nowrap >Start TimeDate:</td>
	<td width="381" align="left" valign="top">
      <DP:RICHDATEPICKER id="dtpStart" runat="server"  Height="24px" LabelText="Please, select a date" Culture="en-GB"  PadSingleDigits="True" DateFormatting="MMMM DD YYYY" NullableLabelText="Select a Dated" ControlDisplay="TextBoxImage" ImageUrl="~/images/calendar.gif"  DateTimeOrdering="TimeThenDate" TimePickerEnabled="True" TimeButtonType="Image">
        <LabelStyle BorderStyle="Dotted" ForeColor="Red" BorderColor="IndianRed" BackColor="White"> </LabelStyle>
      </DP:RICHDATEPICKER>
	  <span class="style3">
	  <asp:CustomValidator id="CustomValidator1" runat="server"
        ControlToValidate="dtpStart"
		ErrorMessage="* start must be earlier than end!"
        OnServerValidate="ServerValidateDateTimes"
        Display="Static"
        >	     </asp:CustomValidator>
	  </span>	 </td>
   
  </tr>
   <tr>
    <td align="left" valign="top" nowrap>End TimeDate:</td>
	<td align="left" valign="top">
      <DP:RICHDATEPICKER id="dtpEnd" runat="server"  Height="24px" LabelText="Please, select a date" Culture="en-GB"  PadSingleDigits="True" DateFormatting="MMMM DD YYYY" NullableLabelText="Select a Dated" ControlDisplay="TextBoxImage" ImageUrl="~/images/calendar.gif"  DateTimeOrdering="TimeThenDate" TimePickerEnabled="True" TimeButtonType="Image">
        <LabelStyle BorderStyle="Dotted" ForeColor="Red" BorderColor="IndianRed" BackColor="White"> </LabelStyle>
      </DP:RICHDATEPICKER>
      <span class="style3">
      <asp:CustomValidator id="CustomValidator2" runat="server"
        ControlToValidate="dtpEnd"
		ErrorMessage="* end must be later than start!"
        OnServerValidate="ServerValidateDateTimes"
        Display="Static"
        > 
		</asp:CustomValidator>
	 
      </span></td>  
   </tr>

</table>
</asp:Placeholder>  
</asp:Panel>
   

 
 <p>&nbsp; </p>
     <p>
      <asp:button id="SubmitBtn" onclick="SubmitBtn_Click" runat="server" text="Submit" />


</form>


</body>
</html>
