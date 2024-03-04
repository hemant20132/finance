<?php
include $_SERVER['DOCUMENT_ROOT'].'/Project/Classes/unique_star_config.php';
if(!class_exists('new_todocontroller')){
    
    class new_todocontroller {   /*creating new class for to do list*/
        function new_todocontroller(){
            //print_r($_GET); die;
            if(isset($_GET['showtasklist'])){
                $this->getTaskList();
            }
            else if(isset($_GET['createlist'])){
                $this->createList();
            }
            else if(isset($_GET['createtasklist'])){
                $this->createTaskList();
            }
            else if(isset($_GET['editlist'])){
                $this->editList();
            }
            else if(isset($_GET['deletelist'])){
                $this->deleteList();
            }
            else if(isset($_GET['markcomplete'])){
                $this->markComplete();
            }
        }
        
        function getListSelectOption(){
            global $unique_star_config;
            $listArray=$unique_star_config->get_results("select * from list");
            if($listArray){
                foreach ($listArray as $list){
                    $option .="<option value='$list->ListID'>$list->ListName</option>";
                }
                return $option;
            }
            return "<option>Sorry! No List name found.</option>";
        }
        
        function getTaskList(){
            global $unique_star_config;
            
			if(isset($_GET['tid']))
			{
				$sql1="select lists.ListID, listtasks.TaskName, listtasks.status 
				from lists, listtasks where lists.ListName='".$_GET['tid']."' 
				and lists.ListID = listtasks.ListID";
                $taskListArray = $unique_star_config->get_results($sql1);
            }
			
			if(!isset($_GET['tid']))
			{
				$sql1="select * from listtasks";
                $taskListArray = $unique_star_config->get_results($sql1);
            }
			
			
            if($taskListArray){
                foreach($taskListArray as $listtask){
                    $status = '';
                    if($listtask->status=="" || $listtask->status==0){
                        $status = "Active";    
                    }elseif($listtask->status==1){
                        $status = "Completed";
                    }
                    
                    $completed = '';
                    if($status=='Active'){
                        $completed = '<a class="" onclick="markComplete('.$listtask->id.')" href="javascript:void(0)">
                        <img src="img/icons/misc/completed.png" title="Mark as complete"></img>
                    </a>';
                    }
                    
                 $tr .='<tr><td>'.$status.'</td><td>'.$listtask->TaskName.'</td>'
                    . '<td>'
                    .'<a class="" onclick="EditList('.$listtask->id.',\''.trim($listtask->TaskName).'\')" href="javascript:void(0)">
                        <img src="img/icons/misc/gridedit.gif"></img>
                    </a>'
                    .'<a class="" onclick="DeleteList('.$listtask->id.')" href="javascript:void(0)">
                        <img src="img/icons/misc/gridtrash.gif"></img>
                    </a>
                   '.$completed.'
                    '
                    .'</td>'
                    .'</tr>';
                }
                
                if(isset($_GET['tid'])){
//                    $tr.="<tr><td colspan='3'><span id='add_new_tasks' class='btn btn-primary pull-right'>Add New Tasks</a></a>";
                   echo json_encode(array('res'=>1,'data'=>$tr,'id'=>$_GET['tid'])); exit;
                }
               echo $tr;
            }
            else 
			{
				if(isset($_GET['tid'])){
//                    $tr.="<tr><td colspan='3'><span id='add_new_tasks' class='btn btn-primary pull-right'>Add New Tasks</a></a>";
                   echo json_encode(array('res'=>1,'data'=>$tr,'id'=>$_GET['tid'])); exit;
                }
				else
				{
					return NULL;
				}
			}
        }
        
        function createList(){
            global $unique_star_config;
            $listname = trim($_GET['ListName']);
            $userid = $_GET['UserID'];
            if(empty($listname)){
                echo json_encode (array('res'=>2,'msg'=>'Please enter name.'));
            }else{
                $unique_star_config->insert("list",array('ListName'=>$listname,'UserID'=>$userid));
                echo json_encode (array('res'=>1,'msg'=>'List has been created successfully..!'));
            }
        }
		
		
		
        function createTaskList(){
            global $unique_star_config;
            $listname = trim($_GET['TaskName']);
            $taskid = $_GET['ListID1'];

			if(empty($listname)){
                echo json_encode (array('res'=>2,'msg'=>'Please enter task name.'));
            }
			else{
                $unique_star_config->insert("listtasks",array('TaskName'=>$listname,'ListID'=>$taskid));
                echo json_encode (array('res'=>1,'msg'=>'Task has been created successfully..!'));
            }
        }
        
        function editList(){
            global $unique_star_config;
            $listname = trim($_GET['ListName']);
            $userid = $_GET['UserID'];
            $listid = $_GET['ListID'];
            if(empty($listname)){
                echo json_encode (array('res'=>2,'msg'=>'Please enter name.'));
            }else{
                $unique_star_config->update("listtasks",array('id'=>$listid), array('TaskName'=>$listname));
                echo json_encode (array('res'=>1,'msg'=>'List has been updated successfully..!'));
            }
        }
        
        function deleteList(){
            global $unique_star_config;
            $listname = trim($_GET['ListName']);
            $userid = $_GET['UserID'];
            $listid = $_GET['ListID'];
            if(empty($listid)){
                echo json_encode (array('res'=>2,'msg'=>'Please select task.'));
            }else{
                $unique_star_config->delete("listtasks",false,' where id='.$listid);
                echo json_encode (array('res'=>1,'msg'=>'List has been deleted successfully..!'));
            }
        }
        
        function markComplete(){
            global $unique_star_config;
            $listid = $_GET['ListID'];
            
                $unique_star_config->update("listtasks",array('id'=>$listid), array('status'=>1));
                echo json_encode (array('res'=>1,'msg'=>'List has been complete marked..!'));
            
        }
    
    }
    
    $new_todocontroller=new new_todocontroller();
}
