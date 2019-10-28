<%
season = Request.QueryString("season")
If season = "" Then season = "3"
season = CInt(season)
%>
<html>
<%If season = 1 Then%>
<body bgcolor="#85B2F6">
<%Else%>
<body bgcolor="#E2D5B9">
<%End If%>
<title>Sac Lunch</title>
<LINK REL=stylesheet HREF="style.css" TYPE="text/css">
<script src="sorttable.js" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" href="subModal/subModal.css" />
<script type="text/javascript" src="subModal/common.js"></script>
<script type="text/javascript" src="subModal/subModal.js"></script>
<%
	'strConn = "PROVIDER=MSDASQL; Driver={SQL Server}; Server=localhost; Database=CP; UID=jais; PWD=jais;" 
	strConn = "PROVIDER=MSDASQL; Driver={SQL Server}; Server=(local); Database=v2000v_com;UID=v2000v; PWD=v2000v_pass;" 
	strConn = "PROVIDER=MSDASQL; Driver={SQL Server}; Server=coreyfields.db.12094096.hostedresource.com; Database=coreyfields;UID=coreyfields; PWD=Alvenu12!!!;" 

	dim RsGame, Rs, RsSum, RsGameStats, RsGameStatsSum, RsStatsToDate, RsAux
	Set Rs = Server.CreateObject("ADODB.Recordset")

	sql = "getStatsAllSeasons"
	Rs.Open Sql, strConn, 3, 3
%>
<img src="images/saclogo2.jpg">&nbsp;&nbsp;&nbsp;
<img src="images/nblogo2.gif">&nbsp;&nbsp;&nbsp;
<img src="images/ThirdSeasonAnniversary.jpg">
<br>
<font class="bodytext">
<a href="index.asp"><< back</a>
<br><br>
<b>Batting Title Winners By Season</b>
<br><br>
2018 Summer: <b>Chris Shrader&nbsp;&nbsp;(.667)</b><br>
2018 Spring: <b>Eric "Tico" Torres&nbsp;&nbsp;(.622)</b><br>
2017 Spring: <b>Matt LeFevre&nbsp;&nbsp;(.720)</b><br>
2016 Summer: <b>Corey Fields&nbsp;&nbsp;(.656)</b><br>
2016 Spring: <b>Eric "Tico" Torres&nbsp;&nbsp;(.763)</b><br>
2015 Summer: <b>Goose&nbsp;&nbsp;(.833)</b><br>
2015 Spring: <b>Goose&nbsp;&nbsp;(.786)</b><br>
2014 Summer: <b>Goose&nbsp;&nbsp;(.735)</b><br>
2014 Spring: <b>Goose&nbsp;&nbsp;(.656)</b><br>
2013 Summer: <b>Aaron Melton&nbsp;&nbsp;(.667 - via tiebreaker)</b><br>
2013 Spring: <b>Matt LeFevre&nbsp;&nbsp;(.641)</b><br>
2012 Summer: <b>Aaron Melton&nbsp;&nbsp;(.714)</b><br>
2012 Spring: <b>Aaron Melton&nbsp;&nbsp;(.635)</b><br>
2011 Summer: <b>Aaron Melton&nbsp;&nbsp;(.667)</b><br>
2011 Spring: <b>Tom McGough&nbsp;&nbsp;(.660)</b><br>
2010 Summer: <b>Aaron Melton&nbsp;&nbsp;(.606)</b><br>
2010 Spring: <b>Boof&nbsp;&nbsp;(.578)</b><br>
2009 Summer: <b>Josh Cook&nbsp;&nbsp;(.718)</b><br>
2009 Spring: <b>Chris Shrader&nbsp;&nbsp;(.641)</b><br>
2008 Summer: <b>Eric "Tico" Torres&nbsp;&nbsp;(.675)</b><br>
2008 Spring: <b>Aaron Melton&nbsp;&nbsp;(.717)</b><br>
2007 Summer: <b>Aaron Melton&nbsp;&nbsp;(.667)</b><br>
2007 Summer: <b>Corey Fields&nbsp;&nbsp;(.703)</b><br>
2007 Spring: <b>Aaron Melton&nbsp;&nbsp;(.612)</b><br>
2006 Summer: <b>Tom McGough&nbsp;&nbsp;(.586)</b><br>
2006 Spring: <b>Ty Travis&nbsp;&nbsp;(.676)</b><br>
<br><br>
Click each column to sort (ascending or descending)
</font>
<br>
<font class="bodytextbold">Total Stats for all seasons</font><!--<br><font class="bodytext">(player must have minimum of 4 games played: "The So-Cal Slugger" Clause)</font--><br>
<table cellpadding="3" cellspacing="1" border="0" width="700" class="sortable" id="1">
	<tr bgcolor="#BBBBBB">
		<td nowrap class="bodytext2" width="50" align="center">Info</td>
		<td nowrap class="bodytext2"><b>Player</b></td>
		<td nowrap class="bodytext2"><b>G</b></td>
		<td nowrap class="bodytext2"><b>AB</b></td>
		<td nowrap class="bodytext2"><b>R</b></td>
		<td nowrap class="bodytext2"><b>RBI</b></td>
		<td nowrap class="bodytext2"><b>SO</b></td>
		<td nowrap class="bodytext2"><b>BB</b></td>
		<td nowrap class="bodytext2"><b>1B</b></td>
		<td nowrap class="bodytext2"><b>2B</b></td>
		<td nowrap class="bodytext2"><b>3B</b></td>
		<td nowrap class="bodytext2"><b>HR</b></td>
		<td nowrap class="bodytext2"><b>H</b></td>
		<td nowrap class="bodytext2"><b>TB</b></td>
		<td class="bodytext2" bgcolor="#969696"><b>Batting Avg.</b></td>
		<td class="bodytext2" bgcolor="#969696"><b>Slugging %</b></td>
		<td class="bodytext2" bgcolor="#969696"><b>On-Base %</b></td>
		<td class="bodytext2" bgcolor="#969696"><b>OPS</b></td>
	</tr>
<%
	'iCount = 0
	Do Until rs.EOF
		'If CInt(rs("G")) > 3 Then
%>
			<tr bgcolor="#eeeeee">
				<td nowrap class="bodytext" align="center"><%If rs("bio") = "1" Then%><img src="images/<%=rs("pic")%>" width="20"><a href="#" onclick="showPopWin('bio.asp?pid=<%=rs("pid")%>', 600, 625, null);">Bio</a><%Else%>Bio<%End If%></td>
				<td nowrap class="bodytext"><%=rs("name")%></td>
				<td nowrap class="bodytext"><%=rs("G")%></td>
				<td nowrap class="bodytext"><%=rs("AB")%></td>
				<td nowrap class="bodytext"><%=rs("R")%></td>
				<td nowrap class="bodytext"><%=rs("RBI")%></td>
				<td nowrap class="bodytext"><%=rs("SO")%></td>
				<td nowrap class="bodytext"><%=rs("BB")%></td>
				<td nowrap class="bodytext"><%=rs("1B")%></td>
				<td nowrap class="bodytext"><%=rs("2B")%></td>
				<td nowrap class="bodytext"><%=rs("3B")%></td>
				<td nowrap class="bodytext"><%=rs("HR")%></td>
				<td nowrap class="bodytext"><%=rs("H")%></td>
				<td nowrap class="bodytext"><%=rs("TB")%></td>
				<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(rs("BA"), 3)%></b></td>
				<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(rs("SLG"), 3)%></b></td>
				<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(rs("OBP"), 3)%></b></td>
				<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(rs("OPS"), 3)%></b></td>
			</tr>
<%
		'End If
		rs.MoveNext
	Loop

	Set RsSum = Server.CreateObject("ADODB.Recordset")

	sql = "getStatsSumAllSeasons"
	RsSum.Open Sql, strConn, 3, 3
%>
</table>
<table border="0" width="700">
	<tr><td colspan="15" class="bodytext">G=Games, AB=At Bats, R=Runs, RBI=Runs Batted In, SO=Strikeouts, BB=Walks, 1B=Singles, 2B=Doubles, 3B=Triples, HR=Home Runs, H=Hits, TB=Tagged Bases, BA=Batting Avg., SLG=Slugging %, OBP=On-Base Percentage, OPS=SLG+OBP</td></tr>
</table>
<table cellpadding="3" cellspacing="1" border="0" width="600">
	<tr bgcolor="#BBBBBB">
		<td nowrap class="bodytext2" width="180" align="center" align="right"><b>Team Stats as of <%=FormatDateTime(Now(), 2)%></b></td>
		<td nowrap class="bodytext2">AB</td>
		<td nowrap class="bodytext2">R</td>
		<td nowrap class="bodytext2">RBI</td>
		<td nowrap class="bodytext2">SO</td>
		<td nowrap class="bodytext2">BB</td>
		<td nowrap class="bodytext2">1B</td>
		<td nowrap class="bodytext2">2B</td>
		<td nowrap class="bodytext2">3B</td>
		<td nowrap class="bodytext2">HR</td>
		<td nowrap class="bodytext2">H</td>
		<td nowrap class="bodytext2">TB</td>
		<td class="bodytext2" bgcolor="#969696"><b>Batting Avg.</b></td>
		<td class="bodytext2" bgcolor="#969696"><b>Slugging %</b></td>
		<td class="bodytext2" bgcolor="#969696"><b>On-Base %</b></td>
		<td class="bodytext2" bgcolor="#969696"><b>OPS</b></td>
	</tr>
	<tr>
		<td class="bodytext" bgcolor="#BBBBBB"><!--(includes all players, regardless of games played)--></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("AB")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("R")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("RBI")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("SO")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("BB")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("1B")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("2B")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("3B")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("HR")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("H")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsSum("TB")%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><br><%=FormatNumber(RsSum("BA"), 3)%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><br><%=FormatNumber(RsSum("SLG"), 3)%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><br><%=FormatNumber(RsSum("OBP"), 3)%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><br><%=FormatNumber(RsSum("OPS"), 3)%></b></td>
	</tr>
</table>
<%
	Set RsAux = Server.CreateObject("ADODB.Recordset")

	sql = "getAuxStats"
	RsAux.Open Sql, strConn, 3, 3
%>
<br><br>
<font class="bodytextbold">Additional Stats (all seasons)</font><!--<br><font class="bodytext">("The So-Cal Slugger" Clause: Player must have minimum of 4 games played)</font--><br>
<table cellpadding="3" cellspacing="1" border="0" width="700" class="sortable" id="1">
	<tr bgcolor="#BBBBBB">
		<td nowrap class="bodytext2" width="50" align="center">Info</td>
		<td nowrap class="bodytext2"><b>Player</b></td>
		<td nowrap class="bodytext2"><b>G</b></td>
		<td nowrap class="bodytext2"><b>R/G</b></td>
		<td nowrap class="bodytext2"><b>RBI/G</b></td>
		<td nowrap class="bodytext2"><b>SO/G</b></td>
		<td nowrap class="bodytext2"><b>BB/G</b></td>
		<td nowrap class="bodytext2"><b>H/G</b></td>
		<td nowrap class="bodytext2"><b>TB/G</b></td>
	</tr>
<%
	'iCount = 0
	Do Until rsAux.EOF
		'If CInt(rsAux("G")) > 3 Then
%>
			<tr bgcolor="#eeeeee">
				<td nowrap class="bodytext" align="center"><%If rsAux("bio") = "1" Then%><img src="images/<%=rsAux("pic")%>" width="20"><a href="#" onclick="showPopWin('bio.asp?pid=<%=rsAux("pid")%>', 600, 625, null);">Bio</a><%Else%>Bio<%End If%></td>
				<td nowrap class="bodytext"><%=rsAux("name")%></td>
				<td nowrap class="bodytext"><%=rsAux("G")%></td>
				<td nowrap class="bodytext"><%=FormatNumber(rsAux("RPG"), 3)%></td>
				<td nowrap class="bodytext"><%=FormatNumber(rsAux("RBIPG"), 3)%></td>
				<td nowrap class="bodytext"><%=FormatNumber(rsAux("SOPG"), 3)%></td>
				<td nowrap class="bodytext"><%=FormatNumber(rsAux("BBPG"), 3)%></td>
				<td nowrap class="bodytext"><%=FormatNumber(rsAux("HPG"), 3)%></td>
				<td nowrap class="bodytext"><%=FormatNumber(rsAux("TBPG"), 3)%></td>
			</tr>
<%
		'End If
		rsAux.MoveNext
	Loop

	Set RsSum = Server.CreateObject("ADODB.Recordset")

	sql = "getStatsSumAllSeasons"
	RsSum.Open Sql, strConn, 3, 3
%>
</table>
<!-- add key -->
</body>
</html>

<%
	Set rs = Nothing
	Set rsGame = Nothing
	Set rsGameStats = Nothing
	Set RsBa = Nothing
	Set RsSum = Nothing
	Set RsGameStatsSum = Nothing
	Set RsStatsToDate = Nothing
	Set rsAux = Nothing
%>