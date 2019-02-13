import cympy
import xlrd
import xlsxwriter
import cympy.db
conn_info = cympy.db.ConnectionInformation()

conn_info.Network = cympy.db.OracleDataSource()
#Enter Service name
conn_info.Network.ServiceName = "Database Name "
#Enter CYME Network SchemaName
conn_info.Network.SchemaName = "SchemaName"
conn_info.Network.UserID = "User ID"
conn_info.Network.Password = "Password"

conn_info.Equipment = cympy.db.MDBDataSource()
conn_info.Equipment.Path = r"Equiment Database path"

cympy.db.Connect(conn_info)
print("Connected to database")
listOfNetworks = cympy.db.ListNetworks()
#Open sheet and import feeders to study
workbook_study = xlrd.open_workbook('Folder Location\Test.xlsx')
f_worksheet = workbook_study.sheet_by_index(0)
f_list=[]
for i in range(1,f_worksheet.nrows):
	f_list.append(f_worksheet.cell_value(i,0))
#Start study for feeders in f_list
Radial_list=[]
for study_f in f_list:
	cympy.study.New()
	cympy.study.LoadNetwork(study_f, cympy.enums.LoadNetworkOption.AllDependencies, 1)
	# Get Switches and Sectionalizer
	l_switch = cympy.study.ListDevices(cympy.enums.DeviceType.Switch)
	l_sectionalizer = cympy.study.ListDevices(cympy.enums.DeviceType.Sectionalizer)	
	# Creation of report Switch Status
	for sw in l_switch:
		sw.SetValue("ABC", "ClosedPhase")
		#print (sw.DeviceNumber,"ABC phases Closed")
	for sec in l_sectionalizer:
		sec.SetValue("ABC", "ClosedPhase")
	#Create a of list of spot loads
	load_list = cympy.study.ListDevices(cympy.enums.DeviceType.SpotLoad)
	#Create a list of overhead secions
	Overhead_list = cympy.study.ListDevices(cympy.enums.DeviceType.OverheadByPhase)
	#Filter overhead list for study feeder only
	NOverhead_list=[]
	for feeder in Overhead_list:
		if feeder.NetworkID==study_f:
			NOverhead_list.append(feeder)	
	for sec_over in NOverhead_list:
		#Disconnect overhead section
		sec_over.SetValue("Disconnected", "ConnectionStatus")
		#Create list of Isolated loads
		isL = []	
		for i in load_list:
			if len((cympy.study.QueryInfoDevice("IsIsolated",i.DeviceNumber, cympy.enums.DeviceType.SpotLoad)))> 2:
				isL.append(int(cympy.study.QueryInfoDevice("DwCustT",i.DeviceNumber,cympy.enums.DeviceType.SpotLoad)))			
		if sum(isL)>0:
			Radial_list.append([str(sec_over.NetworkID),(sec_over.DeviceNumber),sum(isL),'Overhead'])		
		#Reconnect overhead section
		sec_over.SetValue("Connected", "ConnectionStatus")
	cympy.study.Close(False)

workbook = xlsxwriter.Workbook("Result Folder\Overhead Radials.xlsx")
worksheet = workbook.add_worksheet("Radial_list")
worksheet.write(0,0,"Feeder Number")	
worksheet.write(0,1,"Section")
worksheet.write(0,2,"No. of Customers")
worksheet.write(0,3,"Type")
j=1
for feeder,section,load,type in (Radial_list):
	worksheet.write(j,0,feeder)
	worksheet.write(j,1,section)
	worksheet.write(j,2,load)
	worksheet.write(j,3,type)
	j+=1
workbook.close()

