<%
season = Request.QueryString("season")
If season = "" Then
%>
    <!-- #include file="splash.asp" -->
<%
    Response.End
End If
season = CInt(season)
%>
<html>
<%If season = 1 Then%>
<body bgcolor="#85B2F6">
<%ElseIf season = 5 Then%>
<body bgcolor="#FEABD7">
<%ElseIf season = 2 or season = 3 or season = 4 or season = 9 or season = 10 or season = 11 or season = 12 or season = 13 or season = 14 or season = 15 or season = 16 or season = 17 or season = 18 or season = 19 or season = 20 or season = 21 or season = 22 Then%>
<body bgcolor="#E2D5B9">
<%ElseIf season = 6 or season = 7 Then%>
<body bgcolor="#D6D6D6">
<%ElseIf season = 8 Then%>
<body bgcolor="#88B536">
<%End If%>
<title>Sac Lunch</title>
<LINK REL=stylesheet HREF="style.css" TYPE="text/css">
<script src="sorttable.js" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" href="subModal/subModal.css" />
<script type="text/javascript" src="subModal/common.js"></script>
<script type="text/javascript" src="subModal/subModal.js"></script>

<table width="800"><tr><td valign="top">
<%If season = 1 Then%>
<img src="images/logo.jpg">
<%ElseIf season = 2 Then%>
<img src="images/saclogo.jpg"></td>
<td valign="top" class="bodytext"><font size="6"><b>"BALLS OUT, EVERY INNING"</b>
<br><br>
</font>
<font size="3"><b>
<a href="http://www.usssa.com/sports/Tournament3.asp?TournamentID=245688" target="new">Oktoberfest tournament results</a>
</b>
<br>
</font>
<font size="2"><b>
The Jackalopes<br>
Result: 5th place finish (12 teams)
</font>
<%ElseIf season = 5 Then%>
<img src="images/nblogo2.gif">
<!--td valign="top"><img src="images/EricNBFinal.gif"></td-->
<td valign="top"><img src="images/weirdguy.jpg"></td>
<%ElseIf season = 3 Then%>
<img src="images/saclogo.jpg"></td>
<td valign="top"><img src="images/rasheed.jpg" width="200" height="300"></td>
<td valign="top" class="bodytext"><font size="6"><b>"BALLS OUT, EVERY INNING"</b>
<br><br>
<a href="images/DSC01534.jpg" target="11"><img src="images/DSC01534-sm.jpg" border="0"></a>
<a href="images/DSC01540.jpg" target="11"><img src="images/DSC01540-sm.jpg" border="0"></a>
<a href="images/DSC01541.jpg" target="11"><img src="images/DSC01541-sm.jpg" border="0"></a>
<a href="images/DSC01543.jpg" target="11"><img src="images/DSC01543-sm.jpg" border="0"></a>
</td>
<%ElseIf season = 6 or season = 7 Then%>
<img src="images/NewSacLogo2.gif" border="4">
<%ElseIf season = 8 Then%>
<img src="images/saclunch2009-horizontal.jpg" border="4">
<%ElseIf season = 9 or season = 11 or season = 12 or season = 13 or season = 14 or season = 15 or season = 16 or season = 17 or season = 18 or season = 19 or season = 20 or season = 21 or season = 22 or season = 23 or season = 24 or season = 25 or season = 26 Then%>
<img src="images/saclogonew.gif">
<img src="images/weirdguy.jpg">
<%Else%>
<img src="images/EricSacFinal.gif"></td>
<%End If%>
</td>
<td align="right" valign="top" class="bodytext">
<!--font class="bodytext"><b>seasons:</b></font><br-->
<a href="index.asp"><b><< back to homepage</b></a>
<br><br>
<a href="index.asp?season=26"><b>2018 summer</b></a>
<br><br>
<a href="index.asp?season=25"><b>2018 spring</b></a>
<br>
<a href="index.asp?season=24"><b>2017 spring</b></a>
<br>
<a href="index.asp?season=23"><b>2016 summer</b></a>
<br>
<a href="index.asp?season=22"><b>2016 spring</b></a>
<br>
<a href="index.asp?season=21"><b>2015 summer</b></a>
<br>
<a href="index.asp?season=20"><b>2015 spring</b></a>
<br>
<a href="index.asp?season=19"><b>2014 summer</b></a>
<br>
<a href="index.asp?season=18"><b>2014 spring</b></a>
<br>
<a href="index.asp?season=17"><b>2013 summer</b></a>
<br>
<a href="index.asp?season=16"><b>2013 spring</b></a>
<br>
<a href="index.asp?season=15"><b>2012 summer</b></a>
<br>
<a href="index.asp?season=14"><b>2012 spring</b></a>
<br>
<a href="index.asp?season=13"><b>2011 summer</b></a>
<br>
<a href="index.asp?season=12"><b>2011 spring</b></a>
<br>
<a href="index.asp?season=11"><b>2010 summer</b></a>
<br>
<a href="index.asp?season=10"><b>2010 spring</b></a>
<br>
<a href="index.asp?season=9"><b>2009 summer</b></a>
<br>
<a href="index.asp?season=8"><b>2009 spring</b></a>
<br>
<a href="index.asp?season=7"><b>2008 summer</b></a>
<br>
<a href="index.asp?season=6"><b>2008 spring</b></a>
<br>
<a href="index.asp?season=4"><b>2007 summer</b></a>
<br>
<a href="index.asp?season=5"><b>2007 nut brunch</b></a>
<br>
<a href="index.asp?season=3"><b>2007 spring</b></a>
<br>
<a href="index.asp?season=2"><b>2006 summer</b></a>
<br>
<a href="index.asp?season=1"><b>2006 spring</b></a>
<br><br>
<a href="totals.asp"><b>Lifetime Stats</b></a>
<br>
<%If season = 1 Then%>
<a href="images/schedule.jpg" target="1"><b>Schedule</b></a>
<%ElseIf season = 2 Then%>
<a href="sacschedule.txt" target="1"><b>Schedule</b></a>
<%ElseIf season = 3 Then%>
<a href="sacschedule.jpg" target="1"><b>Schedule</b></a>
<%ElseIf season = 4 Then%>
<a href="sacschedulesummer07.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 5 Then%>
<a href="nbschedulesummer07.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 6 Then%>
<a href="sacschedulespring08.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 7 Then%>
<a href="sacschedulesummer08.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 8 Then%>
<a href="sacschedulespring09.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 9 Then%>
<a href="sacschedulesummer09.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 10 Then%>
<a href="sacschedulespring10.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 11 Then%>
<a href="sacschedulesummer10.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 12 Then%>
<a href="sacschedulespring11.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 13 Then%>
<a href="sacschedulesummer11.jpg" target="1"><b>Schedule</b></a>
<%Elseif season = 14 Then%>
<a href="sacschedulespring12.jpg" target="1"><b>Schedule</b></a>
<%Else%>
<a href="http://www.quickscores.com/Orgs/ResultsDisplay.php?OrgDir=ferndalemi&LeagueID=949633" target="1"><b>Schedule & Standings</b></a>
<%End If%>
<!--br>
<a href="http://libertyparkofamerica.com" target="2"><b>League Standings</b></a>
<br>
<a href="http://www.libertyparkofamerica.com" target="3"><b>Liberty Fuckin' Park</b></a-->
<br>
<a href="splash2.asp" target="4"><b>E-League Championship Photos</b></a>
<!--br>
<a href="http://www.coreyfields.com" target="5"><b>coreyfields.com</b></a-->
</td>
</tr>
</table>
<%If season = 6 Then%>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="#C90906">NEWS<br>
8/4/08: Summer season begins this Sunday, August 10th. Game Time announced on Friday. PAYMENT IS $85 AND IS DUE THIS SUNDAY.<br>
<br>
7/28/08: Barry McGough (ejected in the 7th inning on 7/27) to face possible suspension and/or $10,000 fine from Liberty Park. Further details in forthcoming press conference.<br>
</font>
</td></tr>
<tr><td><br></td></tr>
</table>
<%ElseIf season = 7 Then%>
<br>
<!--table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="#C90906">NEWS<br>
8/12/08: Aaron Melton questionable for Sunday's game with freak hamstring injury<br>
</font>
</td></tr>
<tr><td><br></td></tr>
</table-->
<%ElseIf season = 5 Then%>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="#000000">NEWS 8/13/07: Eric "Tico" Torres out for season with surgically-repaired broken thumb.</font>
</td></tr>
<tr><td><br></td></tr>
</table>
<%ElseIf season = 8 Then%>
<br>
<!--table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="#BC5C02">NEWS<br>
Rain makeup game August 9th, 5:00 on Diamond 2<br><br>
</font>
</td></tr>
<tr><td></td></tr>
</table-->
<%ElseIf season = 9 Then%>
<br>
<!--table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="#BC5C02">NEWS<br>
Rain makeup game August 9th, 5:00 on Diamond 2<br><br>
</font>
</td></tr>
<tr><td></td></tr>
</table-->
<%ElseIf season = 10 Then%>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="#BC5C02">NEWS 5/12/10: Eric "Tico" Torres out for season with fractured patella<br><br>
</font>
</td></tr>
<tr><td></td></tr>
</table>
<%ElseIf season = 12 Then%>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="red">NEWS 7/12/11: Brad Abitheira's trade request to Kentucky Minor League affiliate DENIED by GM Frank Marinara<br><br>
</font>
</td></tr>
<tr><td></td></tr>
</table>
<%ElseIf season = 15 Then%>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="red">NEWS 9/11/12: In the 4th inning Sunday night, Brad Abitheira singled to mark the Sac Lunch franchise's 3000th base hit!<br><br>
</font>
</td></tr>
<tr><td></td></tr>
</table>
<%ElseIf season = 19 Then%>
<br>
<table border="0" cellpadding="0" cellspacing="0" width="800" class="bodytext">
<tr><td class="bodytextbold">
<font color="red">NEWS 10/19/14: In the 3rd inning Sunday night, Brad Abitheira singled to mark the Sac Lunch franchise's 4000th base hit!<br><br>
</font>
</td></tr>
<tr><td></td></tr>
</table>
<%End If%>
<%
	'strConn = "PROVIDER=MSDASQL; Driver={SQL Server}; Server=192.168.0.105; Database=CP; UID=jais; PWD=jais;" 
	'strConn = "PROVIDER=MSDASQL; Driver={SQL Server}; Server=(local); Database=v2000v_com;UID=v2000v; PWD=v2000v_pass;" 
	strConn = "PROVIDER=MSDASQL; Driver={SQL Server}; Server=coreyfields.db.12094096.hostedresource.com; Database=coreyfields;UID=coreyfields; PWD=Alvenu12!!!;" 

	dim RsGame, Rs, RsSum, RsGameStats, RsGameStatsSum, RsStatsToDate, RsRecap
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set RsRecap = Server.CreateObject("ADODB.Recordset")
	dim gamedate

	'dim rsUsage
	'If Request.ServerVariables("REMOTE_ADDR") <> "71.144.90.252" Then
	'	Set rsUsage = Server.CreateObject("ADODB.Recordset")
	'	sql = "Insert Into usage (ip) VALUES ('" & Request.ServerVariables("REMOTE_ADDR") & "')"
	'	rsUsage.Open Sql, strConn, 3, 3
	'	Set rsUsage = Nothing
	'End If

	sql = "getStats @season=" & season
	Rs.Open Sql, strConn, 3, 3

	sql = "getLastRecap @season=" & season
	RsRecap.Open Sql, strConn, 3, 3
	
	If Not RsRecap.EOF Then
		recap = RsRecap("recap")
		gamedate = RsRecap("gamedate")
		If Right(gamedate, 1) = "-" Then
			gamedate = Left(gamedate, 6)
		ElseIf Right(gamedate, 1) = " " Then
			gamedate = Left(gamedate, 7)
		End If
	End If
%>
<%If recap > "" Then%>
<br>
<font class="bodytextbold">
<b><%=gamedate%> Game recap<%=recap%>
<%If season = 4 Then%><br><br>By all accounts it was a tumultuous season. 3 heartbreaking losses proved to be the fucking nails in the coffin. While they came up short of the championship, Manager Corey Fields was quick to point out that the team has steadily improved by sporting a better record each season. In light of this, USSSA Analysts have placed Sac Lunch as the <b>overwhelming favorite</b> to win the goddamn division in '08. Manager Fields has demanded the team <b>abstain from drinking and smoking</b> during the entire offseason, and each player must report to training camp in March in the best goddamn shape of their lives.
<br><br>
</font>
<%ElseIf season = 6 Then%><br><br>
Some highlights from the season:<br><br>
<li>Standout infielder Tom McGough selflessly agreeing to play 2B after returning from knee surgery, forming the greatest SS-2B combo in all of Men's E with the power-hitting Boof.<br>
<li>Aaron Melton's record breaking year at the plate. Nearly every record was smashed, including Corey Fields' single season batting average record of .703 (Nut Brunch). Aaron started strong and never let up.<br>
<li>Eric "Tico" Torres' return from an injury-plagued year in 2007, standing out again at 3B and performing flawlessly as a backup Pitcher.<br>
<li>The resurgence of Joe Q at the plate, giving the end of the lineup a new look.<br>
<li>Barry McGough's ejection vs. Amp'd Up, possibly the most infamous incident in Sac Lunch history.<br>
<li>Corey Fields' successful battle against his demons, allowing him to focus on his new Zen-style philosophy of pitching.<br>
<li>Ending the season with 6 regulars with at least a .600 batting average<br>
<li>Although Manager Corey Fields was disappointed at his team giving up 108 runs, they did manage to score a team record 191.<br><br>

2008 was truly a year of destiny. Sac Lunch solidified their reputation as the most feared and hated team in Men's E - a true black eye on the face of Liberty. Yes, this team is truly a force of good and evil, chaos and order... and now they take their own unique brand of softball to the next level. Following a press conference meltdown, Manager Corey Fields offered the following chilling statement: "Men's D will never know what hit them. Those fuckers are going down. All of them."
<br><br>

<%ElseIf season = 7 Then%><br><br>
All things considered, Sac Lunch's first season in D-League was called "acceptable" by Manager/Pitcher Corey Fields. At times they struggled against stronger pitching - leading to some of the highest strikeout numbers in Sac Lunch history. They also struggled with errors in the field, sometimes unable to handle the harder-hit balls. These things are expected to improve in the upcoming 2009 campaign with a rigorous offseason training program, which Mr. Fields claims is mandatory for all players.
<br><br>
Let us not forget the E-League championship run, in which all things seemed to go Sac Lunch's way for the entire season. That season alone was one of the most dominating performances by an E-League team, <i>ever</i>. As the sun sets on 2008, the team will face possible challenges. The rumor mill surrounding LCF Aaron Melton's continued negotiations with Biondo's Market will not let up, despite denials from both sides. And longtime RF Joe Q has been threatening retirement for weeks now - his final decision reportedly to come in February.
<br><br>


<%ElseIf season = 8 Then%><br><br>
This season was yet another step in the direction of C-Minor Baseball Pants League. Despite a sputtering offense, Sac Lunch earned an 8-6 record, their first winning record in D-League. In fact, if a few close games had gone their way, Sac Lunch could have been crowned D-League champs. Not surprisingly, the team has decided to switch back to the classic yellow jerseys for summer season.
<br><br>
And of course, this season recap would not be complete without paying tribute to one of the most controversial and excitable players to ever don a Sac Lunch uniform. Joe Q brought a swagger and attitude often missing from this group of even-keel players. Rumors on the street were that GM Frank Marinara had him forced into retirement, most likely due to lingering issues over the infamous 2007 barfight at the Dugout Bar&Grille. Not to mention Liberty Park officials had been pressuring Marinara to "solve" the Joe Q problem. However, his legacy will never be forgotten. The memories are everwhere: disemboweling the trashcan in nearly every dugout, challenging the opposing team to a fight on many occasions, challenging Aaron to a fight in 2007, challenging Corey to a fight every other game, psychotic screams coming from right field...
<br><br>
Joe Q, you are an officer <i>and a gentleman</i>.
<br><br>
<img src="images/DSC00616.JPG" border="4">
<br><br>

<%ElseIf season = 9 Then%><br><br>
Nonetheless, congratulations are in order as the season concluded with yet another winning record in D-League. As the Sacs have continued to slowly improve, 2010 becomes a crucial year; nothing short of a championship ring will be tolerated by team executives.
<br><br>
<%ElseIf season = 11 Then%><br><br>
For two years - since the ascension to D-League - Sac Lunch has been stuck in mediocrity, playing .500 ball season after season. What will it take to reach the next level? This is the difficult question Manager Corey Fields will struggle with during the offseason. Signs of improvement are clear: the infield, outfield, and pitching rotation continue to rank in the Top 10 of most major statistical categories at Liberty Park. However, the bats remain lethargic. Can defense really win championships?
<br><br>
The greatest accomplishment was the addition of the feared and hideous "Strikeout Blouse". This sinister rag, usually billowing on the dugout fence, was only forced on a few regular players. Consider this astounding statistic: Strikeouts went from .929/game to .455/game, a 51% reduction.<br><br>
<br>
<%Else%>
<br><br>
<%End If%>

<%End If%>
<font class="bodytextbold">
<b>
<%If season = 1 Then%>
Final Record: 2-12 (.167)
<%ElseIf season = 2 Then%>
Final Record: 3-8 (.273)
<%ElseIf season = 3 Then%>
Final Record: 10-4 (.714)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3rd Place Finish
<br><br>
Runs Scored: 155, Runs Against: 118
<%ElseIf season = 4 Then%>
<font color="#C90906">Final Record: 8-3 (.727)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3rd Place (3-way tie)</font>
<br>
Runs Scored: 146, Runs Against: 82
<%ElseIf season = 5 Then%>
<font color="blue">Final Record: 6-5 (.545)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;5th Place</font>
<br>
Runs Scored: 140, Runs Against: 149
<%ElseIf season = 6 Then%>
<font color="#C90906">Final Record: 12-2 (.857)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish</font>
<br>
Runs Scored: 191, Runs Against: 108
<%ElseIf season = 8 Then%>
<font color="#BC5C02">Final Record: 8-6 (.571)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish--></font>
<br>
Runs Scored: 133, Runs Against: 127
<%ElseIf season = 9 Then%>
Final Record: 6-5 (.545)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;4th Place Finish (3-way tie)
<br>
Runs Scored: 121, Runs Against: 88
<%ElseIf season = 10 Then%>
Final Record: 6-8 (.429)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6th Place Finish
<br>
Runs Scored: 106, Runs Against: 144
<%ElseIf season = 11 Then%>
Record: 5-6 (.455)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6th Place Finish-->
<br>
Runs Scored: 100, Runs Against: 133
<%ElseIf season = 12 Then%>
Final Record: 3-11 (.214)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6th Place Finish-->
<br>
Runs Scored: 119, Runs Against: 207
<%ElseIf season = 13 Then%>
Final Record: 4-7 (.364)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6th Place Finish-->
<br>
Runs Scored: 121, Runs Against: 142
<%ElseIf season = 14 Then%>
Final Record: 10-4 (.714)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2nd Place Finish
<br>
Runs Scored: 207, Runs Against: 165
<%ElseIf season = 15 Then%>
Final Record: 7-4 (.636)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2nd Place Finish
<br>
Runs Scored: 128, Runs Against: 92
<%ElseIf season = 16 Then%>
Final Record: 12-1-1 (.893)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish
<br>
Runs Scored: 185, Runs Against: 132
<%ElseIf season = 17 Then%>
Final Record: 3-6-1 (.350)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6th Place Finish
<br>
Runs Scored: 131, Runs Against: 119
<%ElseIf season = 18 Then%>
Final Record: 5-8-1 (.393)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6th Place Finish
<br>
Runs Scored: 149, Runs Against: 149
<%ElseIf season = 19 Then%>
Final Record: 6-5 (.545)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;6th Place Finish-->
<br>
Runs Scored: 127, Runs Against: 107
<%ElseIf season = 20 Then%>
Final Record: 12-2 (.857)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish
<br>
Runs Scored: 192, Runs Against: 108
<%ElseIf season = 21 Then%>
Record: 4-7 (.364)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish-->
<br>
Runs Scored: 143, Runs Against: 135
<%ElseIf season = 22 Then%>
Record: 11-3 (.786)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish
<br>
Runs Scored: 204, Runs Against: 135
<%ElseIf season = 23 Then%>
Record: 4-7 (.364)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish-->
<br>
Runs Scored: 139, Runs Against: 162
<%ElseIf season = 24 Then%>
Record: 4-8 (.333)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish-->
<br>
Runs Scored: 117, Runs Against: 171
<%ElseIf season = 25 Then%>
Record: 5-7 (.417)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish-->
<br>
Runs Scored: 112, Runs Against: 187
<%ElseIf season = 26 Then%>
Record: 0-8 (.000)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place Finish-->
<br>
Runs Scored: 43, Runs Against: 119
<%Else%>
<br>
<font color="#C90906">Record: 5-6 (.455)<!--&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1st Place (tied)--></font>
<br><br>
Runs Scored: 107, Runs Against: 129
<%End If%>
</b></font>
<%If season = 1 Then%>
<br><br>(some game stats unavailable due to arson)
<%End If%>
</font>
<br><br>
<font class="bodytextbold"><b>Season Stats</b></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class="bodytext">Click each column to sort (ascending or descending)</font>
<br>
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
		'If iCount mod 2 = 0 Then
		'	bgcolor="#eeeeee"
		'Else
		'	bgcolor="#ffffff"
		'End If
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
			<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><%=FormatNumber(rs("BA"), 3)%></b></td>
			<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><%=FormatNumber(rs("SLG"), 3)%></b></td>
			<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><%=FormatNumber(rs("OBP"), 3)%></b></td>
			<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><%=FormatNumber(rs("OPS"), 3)%></b></td>
		</tr>
<%
		'iCount = iCount + 1
		rs.MoveNext
	Loop

	Set RsSum = Server.CreateObject("ADODB.Recordset")

	sql = "getStatsSum @season=" & season
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
		<td nowrap class="bodytext" bgcolor="#BBBBBB"></td>
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
	Set RsGame = Server.CreateObject("ADODB.Recordset")
	Set RsGameStats = Server.CreateObject("ADODB.Recordset")
	Set RsGameStatsSum = Server.CreateObject("ADODB.Recordset")
	Set RsStatsToDate = Server.CreateObject("ADODB.Recordset")
	Set RsBa = Server.CreateObject("ADODB.Recordset")

	sql = "Select Distinct GameId From gamestat Where season = " & season & " Order By GameId DESC"
	RsGame.Open sql, strConn, 3, 3
	
	Do Until rsGame.EOF
		sql = "getGameStats @GameId = " & rsGame("GameId") & ", @season = " & season
		'response.write(sql)
		'response.end
		If rsGameStats.State = 1 Then rsGameStats.Close
		RsGameStats.Open Sql, strConn, 3, 3
%>
		<br><br>
		<font class="bodytextbold"><b>Game <%=rsGame("GameId")%></font><font class="bodytextbold"> - <%=rsGameStats("vs")%></font></b><br>
		<%If Not IsNull(rsGameStats("recap")) Then%><font class="bodytextbold">Recap<%=rsGameStats("recap")%>
		</font><%End If%>
		<table cellpadding="3" cellspacing="1" border="0" width="700" class="sortable" id="1">
			<tr bgcolor="#BBBBBB">
				<td nowrap class="bodytext2"><b>Player</b></td>
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
				<td class="bodytext2"><b>New Batting Avg.</b></td>
				<td class="bodytext2"><b>New Slugging %</b></td>
				<td class="bodytext2"><b>New On-Base %</b></td>
				<td class="bodytext2"><b>New OPS</b></td>
			</tr>
		<%
			Do Until RsGameStats.EOF
				sql = "getBA @GameId = " & rsGame("GameId") & ", @pid = " & rsGameStats("pid") & ", @season = " & season
				If RsBA.State = 1 Then RsBA.Close
				RsBA.Open Sql, strConn, 3, 3
		%>
				<tr bgcolor="#eeeeee">
					<td nowrap class="bodytext"><%If RsGameStats("bio") = "1" Then%><a href="bio.asp?season=<%=Request.QueryString("season")%>&pid=<%=RsGameStats("pid")%>"><%=RsGameStats("name")%></a><%Else%><%=RsGameStats("name")%><%End If%></td>
					<td nowrap class="bodytext"><%=RsGameStats("AB")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("R")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("RBI")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("SO")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("BB")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("1B")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("2B")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("3B")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("HR")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("H")%></td>
					<td nowrap class="bodytext"><%=RsGameStats("TB")%></td>
					<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(RsGameStats("BA"), 3)%></b></td>
					<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(RsGameStats("SLG"), 3)%></b></td>
					<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(RsGameStats("OBP"), 3)%></b></td>
					<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(RsGameStats("OPS"), 3)%></b></td>
					<td align="center" nowrap class="bodytext"><%=FormatNumber(RsBA("BA"), 3)%></td>
					<td align="center" nowrap class="bodytext"><%=FormatNumber(RsBA("SLG"), 3)%></td>
					<td align="center" nowrap class="bodytext"><%=FormatNumber(RsBA("OBP"), 3)%></td>
					<td align="center" nowrap class="bodytext"><%=FormatNumber(RsBA("OPS"), 3)%></td>
				</tr>
		<%
				RsGameStats.MoveNext
			Loop

		sql = "getGameStatsSum @GameId=" & rsGame("GameId") & ", @season=" & season
		If RsGameStatsSum.State = 1 Then RsGameStatsSum.Close
		RsGameStatsSum.Open Sql, strConn, 3, 3
%>
</table>
<table cellpadding="3" cellspacing="1" border="0" width="600">
	<tr><td></td></tr>
	<tr bgcolor="#BBBBBB">
		<td nowrap class="bodytext2" width="107" align="center" align="right"><b>Team</b></td>
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
		<td nowrap class="bodytext" bgcolor="#BBBBBB"></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("AB")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("R")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("RBI")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("SO")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("BB")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("1B")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("2B")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("3B")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("HR")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("H")%></b></td>
		<td nowrap class="bodytext" bgcolor="#BBBBBB"><b><br><%=RsGameStatsSum("TB")%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><br><%=FormatNumber(RsGameStatsSum("BA"), 3)%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><br><%=FormatNumber(RsGameStatsSum("SLG"), 3)%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><br><%=FormatNumber(RsGameStatsSum("OBP"), 3)%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><br><%=FormatNumber(RsGameStatsSum("OPS"), 3)%></b></td>
	</tr>
<%
		sql = "getStatsToDate @GameId = " & rsGame("GameId") & ", @season = " & season
		If rsStatsToDate.State = 1 Then rsStatsToDate.Close
		rsStatsToDate.Open Sql, strConn, 3, 3
%>
	<tr>
		<td nowrap class="bodytext" bgcolor="#969696"><b>Team Batting Avg. Through Game <%=rsGame("GameId")%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(rsStatsToDate("BA"), 3)%></b></td>
	</tr>
	<tr>
		<td nowrap class="bodytext" bgcolor="#969696"><b>Team Slugging % Through Game <%=rsGame("GameId")%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(rsStatsToDate("SLG"), 3)%></b></td>
	</tr>
	<tr>
		<td nowrap class="bodytext" bgcolor="#969696"><b>Team On-Base % Through Game <%=rsGame("GameId")%></b></td>
		<td nowrap class="bodytext" bgcolor="#969696"><b><%=FormatNumber(rsStatsToDate("OBP"), 3)%></b></td>
	</tr>
</table>


<%		rsGame.MoveNext
	Loop
%>

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
%>