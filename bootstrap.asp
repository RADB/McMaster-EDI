<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Bootstrap 101 Template</title>

    <!-- Bootstrap -->
    <link href="css/bootstrap.min.css" type="text/css" rel="stylesheet">
	<link href="css/bootstrap-toggle.min.css" type="text/css" rel="stylesheet">
	

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
	
	<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="js/jquery-1.11.3.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="js/bootstrap.min.js"></script>
	<script src="js/bootstrap-toggle.min.js"></script>
  </head>
  <body>
<!-- Button trigger modal -->
<button type="button" class="btn btn-primary" data-toggle="modal" radb-query="getHelp.asp?provinceID=1&Section=Demographics&Question=1" data-target="#myModal">Open modal</button>

<img border="0" href="getHelp.asp?provinceID=1&Section=Demographics&Question=1" src="images\Help.png" alt="For Help See Guide" name="Help" title="For Help See Guide" height="20" data-toggle="modal" data-target="#myModal"/>

<!-- Modal -->
<div class="modal fade" id="myModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
  <div class="modal-dialog">
    <div id="modal-content" class="modal-content">
      <!--<div class="modal-header">
        <button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">&times;</span><span class="sr-only">Close</span></button>   
		<h4 class="modal-title" id="myModalLabel"><img border="0" src="images\Help.png" alt="Help" name="Help" title="Help" height="40"/>  Test Modal</h4>
      </div> 
      <div class="modal-body">
        lorem ipsum
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>        
      </div>-->
    </div>
  </div>
</div>

<input type="checkbox" checked data-toggle="toggle" data-off="No" data-on="Yes" data-onstyle="success" data-offstyle="danger" data-size="mini">

	<script language="javascript" type="text/javascript">		
		$('#myModal').on('show.bs.modal', function(e){
			
			$('#modal-content').load( e.relatedTarget.attributes['radb-query'].value, function (r, s, xhr) {
				// this is not needed but you can do other stuff here
				alert('load succeeded');
			});
		});
		
		// this is not needed
		$('#myModal').on('hide.bs.modal', function(){
			alert('test hidden');
		});
	</script>
	
  </body>
</html>