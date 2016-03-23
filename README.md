# HAR-001DataOrganization
Script_DataOrganization
1.	Search the folders with subject ID such 001-01, 001-08, ….  
2.	Visit each folder and find the files that need to be averaged. There is a log.csv which records if one certain file has been calculated or not.
3.	Correct the data format in each file, because for some reasons, the data in several rows at the end does not has the correct separator. 
4.	Run Noam’s Macro. It fetches the TMS and CES threshold from the SubjectThreshold.csv file
5.	Average data by normIntensity, sort by time, normIntensity and channel. Save the average data into the excel spreadsheet.

Script_QueryThreeFiles
1.	Search the folders with subject ID such 001-01, 001-08, ….
2.	Visit each folder and check if F, MEP, CES files for one certain subject and session are all available.
3.	Check if the subject and the session have been analyzed. (There is a log_wholeData.csv to record this information.)
4.	Normalize the amplitudes of MEP and CES by M wave amplitude according to the time and channel
5.	Fetch amplitude and latency of F, MEP and CES and save them into wholeData.csv file
6.	It also checks if the data with the same subject and session is exist in the file. If so, it will update it. This allows us to update exist data if we re-quantify the data.
7.	Fetch the intervention data from SubjectThreshold.csv
8.	Sort the data based on subject, intervention, time and normIntensity
9.	Save the data in the wholeData.csv
