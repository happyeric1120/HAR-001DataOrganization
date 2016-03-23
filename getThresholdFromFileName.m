function [subjectID session TMSRMT CESRMT] = getThresholdFromFileName(fileName, thresholdFileName)
[pathstr name ext] = fileparts(fileName);

nameCell = strsplit(name, ' ');
subjectID = strsplit(nameCell{1}, '-');
subjectID = str2num(subjectID{2});

session = nameCell{2};
session = str2num(session(2));

thresholdData = xlsread(thresholdFileName);
queryRow = thresholdData(find(thresholdData(:,1) == subjectID & thresholdData(:,2) == session), :);

if isempty(queryRow) == 1
    TMSRMT = 0;
    CESRMT = 0;
else
    TMSRMT = queryRow(3);
    CESRMT = queryRow(4);
end


