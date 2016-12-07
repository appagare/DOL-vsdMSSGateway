-- vsMSSGateway_Enable_Disable_queue_script.sql
use vsMSSGateway
go

declare @Disabled tinyint
declare @QueueLabel varchar(20)
-- set this value to the label of the queue being enabled/disabled.
-- select distinct queuelabel from tblQueueList
set @QueueLabel = ''

/*
Toggle the @Disabled value between zero (0) and one (1). 
When zero, the queue is processed normally (i.e. - not disabled). 
When one, messages destined for this queue are ignored (i.e. - standby mode).
*/
set @Disabled = 0

/*
No need to modify anything below here
Step 1 - Update the Disabled bit
*/
update tblQueueList
set Disabled = @Disabled
where QueueLabel = @QueueLabel

/* Step 2 - Show the results */
select * from tblQueueList 
where QueueLabel = @QueueLabel 
order by QueueLabel asc
