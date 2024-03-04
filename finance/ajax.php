	
			$("#finyear").change(function()
			{
		
				$.post({
				url : "companyfinyearcheck.php",
				dataType:'html';
				type: "POST",
				data : {companyname:'$('#company').val()', finyear:'$('#finyear').val()' }
				success: function(data, textStatus, jqXHR)
				{
				alert (data);
				return false;
				},
				error: function (jqXHR, textStatus, errorThrown)
				{
				}

				
				});

			});			
			
	