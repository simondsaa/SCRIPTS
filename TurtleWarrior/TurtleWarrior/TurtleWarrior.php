#!/usr/local/bin/php
<?php 
error_reporting(E_ALL);
ini_set('display_errors', '1');
?>
 <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
 "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head>
  <meta charset="utf-8">

<link rel="stylesheet" type="text/css" href="http://www.cs.sfu.ca/templates/main.css" media="all" />


<link rel="stylesheet" type="text/css" href="../tooltip.css">
  <link href="../css/start/jquery-ui-1.10.3.custom.css" rel="stylesheet">

<meta http-equiv="content-type" content="text/html; charset=ISO-8859-1" />
<meta http-equiv="content-language" content="en" />
<meta name="resource-type" content="document" />
<meta name="robots" content="all" />

<meta name="keywords" content="computing,computer,science,keywords" />
<meta name="description" content="Games made by Behnam Azizi, TurtleWarrior, Computer Science student" />
                                  

		<title>TurtleWarrior - HTML5 - Behnam Azizi</title>
		<link rel="icon" 
		      type="image/png" 
		      href="turtle.png">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>
		<script>
			document.onload = function() {
				document.addEventListener("keydown", function(e) {
					if ([32, 37, 38, 39, 40].indexOf(e.keyCode) > -1) {
						e.preventDefault();
						// Do whatever else you want with the keydown event (i.e. your navigation).
					}
				}, false);

			};
		</script>

		<script src="util.js" type="text/javascript"></script>
		<script src="jquery.hotkeys.js"></script>
		
		<script src="key_status.js"></script>
	
		<script src="collisionHandler.js"></script>
		<script src="Background.js"></script>
		<script src="myFirstGame.js"></script>


		<img src="mainMenu.jpg" id="menu" class="hidden" />
		<img src="cursor_flipped.png" id="cursor2" class="hidden" />
		<img src="cursor.png" id="cursor1" class="hidden" />
		<img src="credits.png" id="credits" class="hidden" />
		<img src="instructions.png" id="instructions" class="hidden" />


		<img src="turtle.png" id="charImage" class="hidden" />
		<img src="turtle_flipped.png" id="charImage2" class="hidden" />
		<img src="bullet.png" id="bulletImage" class="hidden" />
		<img src="bullet_flipped.png" id="bulletImage2" class="hidden" />
		<img src="bubble.png" id="bubbleImage" class="hidden" />

		<!-- Enemies -->
		<img src="enemy_1.png" id="enemy1" class="hidden" />
		<img src="enemy_2.png" id="enemy2" class="hidden" />
		<img src="enemy_3.png" id="enemy3" class="hidden" />
		<img src="enemy_4.png" id="enemy4" class="hidden" />
	

		<!-- Bonuses -->
		<img src="fuel.png" id="bonus1" class="hidden" />
		<img src="ammo.png" id="bonus2" class="hidden" />
		<img src="aid.png" id="bonus3"  class="hidden"/>

		<img src="jetpack.png" id="jetPackImage" class="hidden" />
		<img src="background.gif" id="bgImage" class="hidden" />
		<img src="gameover.png" id="gameOver" class="hidden" />
		<style>
			body {
				font-size: 14pt;
			}
		</style>



</head>

<body>

<!-- section menu -->

<div id="sectionsmenu">
  <table cellpadding="0" cellspacing="0" border="0" align="right">
    <tr>
      <td width="24" id="slc">
        <img src="http://www.cs.sfu.ca/templates/images/pixel.gif" border="0" alt="" height="10"
             width="24" />
      </td>
  <?php
require_once ('../menu.php');
?>
      <td width="24" id="src">
        <img src="http://www.cs.sfu.ca/templates/images/pixel.gif" border="0" alt="" height="10"
             width="24" />
      </td>
      <td width="24">&nbsp;</td>
    </tr>
  </table>
</div>


<div id="wrapperL">

<table id="layout_table" border="0" cellspacing="0" cellpadding="0" width="100%">

<!-- Page Top -->

<tr>
  <td id="top" colspan="3" align="left" valign="top">
    <img height="105" width="380" align="left" alt=""
         src="http://www.cs.sfu.ca/templates/images/top.gif" />
    <img align="right" height="105" width="102" alt=""
         src="http://www.cs.sfu.ca/templates/images/menu-top.gif" />
  </td>
  
  <td class="searchbox" align="right" valign="bottom">
    <form action="http://google.com/u/sfucs" method="get">
      <input name="hl" type="hidden" value="en" />
      <input name="ie" type="hidden" value="UTF-8" />
      <input name="oe" type="hidden" value="utf-8" /> 

      <input name="q" type="text" size="0" id="search" value="search" 
             onfocus="if (this.value == 'search') this.value=' '"
             onblur="if (this.value == '' || this.value == ' ') this.value='search'"
             class="inputtext"
             title="Search SFU domain using Google" />

      <input id="search_btn_reg" type="submit" value="Go" />
    </form>
  </td>

</tr>

<!-- Page Center -->

<tr>
  <td width="60" height="320">
    <img src="http://www.cs.sfu.ca/templates/images/pixel.gif" width="60" border="0" height="320" align="right" alt="" />
  </td> 

  <td id="content" width="100%">&nbsp;


<span id="crumbs"> &raquo; <a href="/" title="back to &quot;Home&quot;">Home</a> &raquo; 
<span title="current section"><?php //echo TITLE ?></span></span>

<!--             -->


     <p> TurtleWarrior is the latest game that I developed using
      	HTML5 canvas. HTML5 canvas is an html equivalent to Adobe Flash.
      	However there are both distadvantages and advantages of using
      	HTML5 Canvas as opposed to flash. One important advantage of HTML5
      	over flash
      	is that the new mobile and portable devices no longer support flash
      	while HTML5 is supported in most new desktop and mobile browsers.
      	A disadvantage of HTML5, however, maybe that it has not yet received
      	enough support that Flash currently has and, due to lack of good documentations 
      	(compared to flash) and powerful libraries for creating animations,
      	it may be a bit harder to make animations and games in HTML5.
      	<h4>Documentations:</h4>
      	<p> I found a couple of good documentations that helped me a lot through development of this game.
      		Here I list a couple of them:
      		<ul>
      			<li><a href="http://www.html5rocks.com/en/tutorials/canvas/notearsgame/" target="_blank">No Tears Guide to HTML5 Games</a></li>
      			<li><a href="http://www.w3schools.com/html/html5_canvas.asp" target="_blank">W3Schools</a></li>
      			<li>And of course, needless to say, StackOverflow.</li>
      			
      		</ul>
      	</p>


		<h4> Interesting facts about the game: </h4>
		<ul>
			<li> The enemies movement is a sinusoidal function. The exact
				function is A*Sin((800-x)/100) + Y0, where:
				<ul>
					<li> Y0: Is a random number between 100 to 500</li>
					<li> A: is the amplitude of the sin(x) which in this case is 100 pixels!
						(i.e., enemies moves between the range of 100 pixels above their initial Y0 position and 100 pixels below)</li>
						<li> X: is the current X position of the enemy on the screen.</li>
					<li> FPS (Frames Per Second): The FPS of this game is set to 24 Frames Per Second. 
						Meaning that each second the screen of the game is completely cleared and drawn 24 times!</li>
				</ul>
				
				</li>
		</ul>
		<div align="center">
			My highest score is 754. What's yours? <br />
			<canvas id="gameCanvas" width="800" height="530" style="border: 2px solid #000000;">TurtleWarrior</canvas>
		</div>


<!--             -->

    </div>
  </td>
  <td id="sidebar_curve" align="right">
    <img width="30" height="315" alt="subtle wave for the sidebar"
         src="http://www.cs.sfu.ca/templates/images/menu-side.gif" />
  </td>
  <td id="sidebar" style="width: 120px">

  </td>
</tr>

<!-- Page Footer -->

<tr>
  <td id="footerL" align="left" valign="bottom">
  </td>
  <td id="footerC" align="left" valign="bottom">
    <div class="footer" style="float: left; margin: 0px 0px 0px 0px; padding-bottom: 20px;"><h5>Webmaster:
      <a href="mailto:csweb@cs.sfu.ca">csweb@cs.sfu.ca</a></h5></div>
  </td>
  <td id="footerCR" align="right">
    <div id="footerCRFG" style="background: transparent url(templates/images/menu-bottom-t.gif) no-repeat right bottom;"></div>
  </td>
  <td id="footerR" align="right" style="width: 120px;">
    <div id="footerRFG" style="width: 120px;"></div>
  </td>
</tr>

</table>

<div class="footer">
      
    <h5>
      <a href="http://fas.sfu.ca/" target="_blank"
         title="external link to: http://fas.sfu.ca">
           Faculty of Applied Sciences</a>
      |
      <a href="http://www.sfu.ca/" target="_blank"
         title="external link to: http://www.sfu.ca">
           Simon Fraser University</a>

    <br /></h5>
</div>

</div>

</body>
</html>

