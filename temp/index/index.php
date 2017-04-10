
<div id="board"> Work! </div>
<?php
/*
	if (!empty($_SERVER['HTTPS']) && ('on' == $_SERVER['HTTPS'])) {
		$uri = 'https://';
	} else {
		$uri = 'http://';
	}
	$uri .= $_SERVER['HTTP_HOST'];
	header('Location: '.$uri.'/xampp/');
	exit;
*/
    //Something is wrong with the XAMPP installation :-(
?>

<style type="text/css">
@import "compass/css3";

* {
  box-sizing:border-box;
  padding:0;
  margin:0;
}

html {
 	height:100%; 
}

body {
 	background:linear-gradient(#adf1ff, #fff); 
  height:100%; 
}

#board {
  position:relative;
  width:200px;
  height:150px;
  margin:40px auto;
  background:linear-gradient(#616161, #474747);
  border:1px solid #9f7037;
  box-shadow:
    inset 0 0 2px 2px rgba(0,0,0,0.3),
    inset 0 30px 30px rgba(255,255,255,0.1);
  text-align:center;
  color:white;
  font-weight:bold;
  font-size:40px;
  font-family:'Monospace';
  letter-spacing:3px;
  text-shadow:0 1px 0 rgba(0,0,0,0.3);
  line-height:150px;
  border-radius:3px;
}

#board:before {
  content:"";
  position:absolute;
  left:-10px;
  right:-10px;
  top:-10px;
  bottom:-10px;
  background:#e8cca0;
  z-index:-1;
  border-radius:0.3em;
  border:#9b6b2f solid 1px;
  box-shadow:
    0 4px 2px 2px rgba(0,0,0,0.2),
    inset 0 1px 0 rgba(255,255,255,0.4),
    inset 0 0 2px rgba(0,0,0,1);
}

#board:after {
  content:"";
  display:block;
  position:absolute;
  border-radius:0.1em 0.1em 0 0.1em;
  width:50px;
  height:8px;
  background:#f1f1f1;
  bottom:-1px;
  right:15px;
  box-shadow:
    inset 0 -4px 1px rgba(0,0,0,0.3),
    -1px -1px 1px rgba(0,0,0,0.2),
    0 2px 0 rgba(0,0,0,0.3);
}
</style>
