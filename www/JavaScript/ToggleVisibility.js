// JavaScript Document - Copyright of Ryan Fletcher - Denude

		// Display Toggle
			function toggle_visibility(id) 
			{
			   var e = document.getElementById(id);
			   if(e.style.display == 'block'){
				  e.style.display = 'none';}
			   else{e.style.display = 'block';}
			}	
		// Drag and Drop Function
			 function allowDrop(us)
			   {us.preventDefault();}
			 function drag(us)
			   {us.dataTransfer.setData("Text",us.target.id);}
			 function drop(us){
				 us.preventDefault();
			 var data=us.dataTransfer.getData("Text");
			 us.target.appendChild(document.getElementById(data));}