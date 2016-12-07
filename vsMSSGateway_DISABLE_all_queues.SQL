-- vsMSSGateway_DISABLE_all_queues.SQL
use vsMSSGateway
go
/*
This script will mark all queues as disabled. This should be run on the standby instance of the vsMSSGateway database unless in disaster recovery mode. The vsMSSGatewayDB service only reads the database upon startup. Therefore, the vsMSSGatewayDB service MUST be recycled after running this script for it to be effective. 

Also, the vsMSSGateway_ENABLE_all_queues.SQL script should be run on the non-standby instance of the vsMSSGateway database and the non-standby vsMSSGatewayDB service should be recycled.
*/
update tblQueueList
set Disabled = 1
/* Step 2 - Show the results */
select * from tblQueueList 
order by QueueLabel asc
go

