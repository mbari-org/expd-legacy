	<%@ Page language="c#"  %>
		<%@ Import Namespace = "System" %>
		<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
		<HTML>
		    <HEAD>
		     <title>WebForm1</title>	
		      <script language=C# runat=server>
			   public void Page_Load(object sender, System.EventArgs e)
			   {	Version vs = Environment.Version;
				Response.Write("ASP.NET version of this Application is :" +vs.ToString());
			   }
		     </script>
		   </HEAD>
		 <body MS_POSITIONING="GridLayout">		
		 </body>
		</HTML>