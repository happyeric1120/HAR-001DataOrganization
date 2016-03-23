clear all
dataPath = '\\R04BRXNAS20.v03.med.va.gov\VHABRXWuY$\IRM\Desktop\Matlab Codes\Data\';
wholeDataSaveFile = [dataPath 'wholeData.xlsx'];
thresholdFileName = [dataPath 'SubjectThreshold.csv'];
log_wholeDataFileName = [dataPath 'log_wholeData.csv'];

% Get all subjects folder
totalSubjectFolders = getTotalSubjectFolders(dataPath);

% Get the log_wholeData.csv to load the subjectIDs that have been analyzed.
processedSubject = {};
if exist(log_wholeDataFileName, 'file') ~= 0
    fLog_wholeData = fopen(log_wholeDataFileName, 'r');
    iteration = 1;
    tline = fgetl(fLog_wholeData);
    while ischar(tline)
        processedSubject{iteration} = tline;
        iteration = iteration + 1;
        tline = fgetl(fLog_wholeData);
    end
    fclose(fLog_wholeData);
end

nonProcessedSubject = {}; % This will save all the subjectID (ex. 15-009) with available F, CES and MEP files
counterNPS = 1; % counter for non processed subject
for folder = totalSubjectFolders
    singleSubjectFolderPath = [dataPath folder{1}];
    singleSubjectData = dir(singleSubjectFolderPath);
    
    availableExcelData = {};
    counterAED = 1; % counter Available Excel Data
    for numName = 1:length(singleSubjectData)
        if singleSubjectData(numName).isdir ~= 1
            [pathStr name ext] = fileparts(singleSubjectData(numName).name);
            if strcmp(ext, '.xlsx') == 1
                availableExcelData{counterAED} = singleSubjectData(numName).name;
                counterAED = counterAED + 1;
            end
        end
    end
                
    
    % 8 sessions
    for session = 2:9
        subjectID = strsplit(folder{1}, '-');
        subjectID = subjectID{2};
        subjectSessionName = ['15-' subjectID ' S' num2str(session)];
        
        fileTypes = {'F', 'CES', 'MEP'};
        checkSum = 0;
        % Check if a subject with session has been analyzed
        % After that, check if F, CES and MEP are all available
        if isempty(cell2mat(strfind(processedSubject, subjectSessionName))) == 1
            for fileType = fileTypes
                subjectSessionNameAndType = [subjectSessionName ' ' fileType{1} '.xlsx'];
                searchFile = strfind(availableExcelData, subjectSessionNameAndType);
                if isempty(cell2mat(searchFile)) ~= 1
                    checkSum = checkSum + 1;
                end
            end
        end
        if checkSum == 3 % all data are available
            nonProcessedSubject{counterNPS} = [singleSubjectFolderPath '\' subjectSessionName];
            counterNPS = counterNPS + 1;
        end
    end
end





for subjectSession = nonProcessedSubject
    [pathStr subjectWithSession ext] = fileparts(subjectSession{1});
    
    
    folderName = pathStr;
    fileName_F_noPath = [subjectWithSession ' F.xlsx'];
    fileName_MEP_noPath = [subjectWithSession ' MEP.xlsx'];
    fileName_CES_noPath = [subjectWithSession ' CES.xlsx'];

    fileName_F = [folderName '\' fileName_F_noPath];
    fileName_MEP = [folderName '\'  fileName_MEP_noPath];
    fileName_CES = [folderName '\' fileName_CES_noPath];

    % Intervention --> Column 5
    thresholdData = xlsread(thresholdFileName);

    % Get subject ID, session (convert to intervention) from filename
    cellFileName = strsplit(fileName_F_noPath, ' ');
    subjectID = strsplit(cellFileName{1}, '-');
    subjectID = str2num(subjectID{2});      % Get subjectID

    session = cellFileName{2};
    session = str2num(session(2));
    intervention = thresholdData(find(thresholdData(:,1) == subjectID & thresholdData(:,2) == session),:);
    intervention = intervention(5);  % Column 5 is intervention







    data_F = xlsread(fileName_F, 'average');
    data_MEP = xlsread(fileName_MEP, 'average');
    data_CES = xlsread(fileName_CES, 'average');

    % Time --> Column 1, NormIntensity --> Column 2, Channel --> Column 3
    % M, MEP, CES Latency --> T1vT2lat (Column 4)
    % AMPWAVE 1 --> Column 13, AMPWAVE 2 --> Column 19
    % F latency --> Column 6

    % The stategy is query the M wave amplitude by getting time and channel
    % from the data_F, data_MEP and dat_CES. Then doing the normalization for
    % the amplitude
    time = unique(data_F(:,1));  % This should output -1, 0, 15, 30
    channel = unique(data_MEP(:,3)); % This should output 0 and 1. Question: how to normalize BP and FCR

    dimension_data_F = size(data_F);
    dimension_data_MEP = size(data_MEP);
    dimension_data_CES = size(data_CES);

    for row = 1:dimension_data_MEP(1)
        temp_data = data_MEP(row, :);
        query_time = temp_data(1);
        query_channel = temp_data(3);

        query_F_result = data_F(find(data_F(:,1) == query_time & data_F(:,3) == query_channel),:);



        % I named these on purpose, so the codes can be used repeatedly.
        latency = temp_data(4);
        amplitude = temp_data(13);
        normIntensity = temp_data(2);

        % If the channel is 0 or 1, then normalize the amplitude
        if ~isempty(query_F_result)
            M_latency = query_F_result(4);
            M_amplitude = query_F_result(13);
            F_latency = query_F_result(6);
            F_amplitude = query_F_result(19);
            amplitude = amplitude/M_amplitude;
        end

        processed_row = [subjectID intervention query_time normIntensity query_channel latency amplitude];

        if row == 1
            saveData = processed_row;
        else
            saveData = cat(1, saveData, processed_row);
        end

    end

    MEP_saveData = saveData;

    for row = 1:dimension_data_CES(1)
        temp_data = data_CES(row, :);
        query_time = temp_data(1);
        query_channel = temp_data(3);

        query_F_result = data_F(find(data_F(:,1) == query_time & data_F(:,3) == query_channel),:);



        % I named these on purpose, so the codes can be used repeatedly.
        latency = temp_data(4);
        amplitude = temp_data(13);
        normIntensity = temp_data(2);

        % If the channel is 0 or 1, then normalize the amplitude
        if ~isempty(query_F_result)
            M_latency = query_F_result(4);
            M_amplitude = query_F_result(13);
            F_latency = query_F_result(6);
            F_amplitude = query_F_result(19);
            amplitude = amplitude/M_amplitude;
        end

        processed_row = [subjectID intervention query_time normIntensity query_channel latency amplitude];

        if row == 1
            saveData = processed_row;
        else
            saveData = cat(1, saveData, processed_row);
        end

    end

    CES_saveData = saveData;

    % Create F_saveData
    F_subjectID_array = zeros(dimension_data_F(1),1)+subjectID;
    F_intervention_array = zeros(dimension_data_F(1),1)+intervention;
    F_saveData = [F_subjectID_array F_intervention_array data_F(:,1:3) data_F(:,4) data_F(:,13) data_F(:,6) data_F(:,19)];


    % Correct the normIntensity
    MEP_saveData(:,4) = round(MEP_saveData(:,4)./5).*5;
    CES_saveData(:,4) = round(CES_saveData(:,4)./5).*5;

    F_header = {'Subject', 'Intervention', 'Time', 'NormIntensity', 'Channel', 'M Latency', 'M Amplitude', 'F Latency', 'F Amplitude'};
    MEP_header = {'Subject', 'Intervention', 'Time', 'NormIntensity', 'Channel', 'MEP Latency', 'MEP Amplitude'};
    CES_header = {'Subject', 'Intervention', 'Time', 'NormIntensity', 'Channel', 'CES Latency', 'CES Amplitude'};
    
    
    % Load the data from excel file and search if the data is in it by
    % subject and session. If so, remove it and use the newer data
    wholeData_F = [];
    wholeData_MEP = [];
    wholeData_CES = [];
    if exist(wholeDataSaveFile, 'file') ~= 0
        wholeData_F = xlsread(wholeDataSaveFile, 'F');
        wholeData_MEP = xlsread(wholeDataSaveFile, 'MEP');
        wholeData_CES = xlsread(wholeDataSaveFile, 'CES');
        
        % Remove data that is existed already
        wholeData_F(wholeData_F(:,1) == subjectID & wholeData_F(:,2) == intervention, :) = [];
        wholeData_MEP(wholeData_MEP(:,1) == subjectID & wholeData_MEP(:,2) == intervention, :) = [];
        wholeData_CES(wholeData_CES(:,1) == subjectID & wholeData_CES(:,2) == intervention, :) = [];
    end
    

    
    wholeData_F = cat(1, wholeData_F, F_saveData);
    wholeData_MEP = cat(1, wholeData_MEP, MEP_saveData);
    wholeData_CES = cat(1, wholeData_CES, CES_saveData);
    
    wholeData_F = sortrows(wholeData_F, [1 2 3 4 5]);
    wholeData_MEP = sortrows(wholeData_MEP, [1 2 3 4 5]);
    wholeData_CES = sortrows(wholeData_CES, [1 2 3 4 5]);
    
    
    xlswrite(wholeDataSaveFile, F_header, 'F', 'A1');
    xlswrite(wholeDataSaveFile, wholeData_F, 'F', 'A2');
    xlswrite(wholeDataSaveFile, MEP_header, 'MEP', 'A1');
    xlswrite(wholeDataSaveFile, wholeData_MEP, 'MEP', 'A2');
    xlswrite(wholeDataSaveFile, CES_header, 'CES', 'A1');
    xlswrite(wholeDataSaveFile, wholeData_CES, 'CES', 'A2');   

%     if exist(wholeDataSaveFile, 'file') == 2
%         F_rowLength = size(xlsread(wholeDataSaveFile, 'F'),1);
%         MEP_rowLength = size(xlsread(wholeDataSaveFile, 'MEP'),1);
%         CES_rowLength = size(xlsread(wholeDataSaveFile, 'CES'),1);
% 
%         xlswrite(wholeDataSaveFile, F_saveData, 'F', ['A' num2str(F_rowLength+2)]);
%         xlswrite(wholeDataSaveFile, MEP_saveData, 'MEP', ['A' num2str(MEP_rowLength+2)]);
%         xlswrite(wholeDataSaveFile, CES_saveData, 'CES', ['A' num2str(CES_rowLength+2)]);
%     else
%         xlswrite(wholeDataSaveFile, F_header, 'F', 'A1');
%         xlswrite(wholeDataSaveFile, F_saveData, 'F', 'A2');
%         xlswrite(wholeDataSaveFile, MEP_header, 'MEP', 'A1');
%         xlswrite(wholeDataSaveFile, MEP_saveData, 'MEP', 'A2');
%         xlswrite(wholeDataSaveFile, CES_header, 'CES', 'A1');
%         xlswrite(wholeDataSaveFile, CES_saveData, 'CES', 'A2');
%     end
    
    % Save the subject with session to log_wholeData.csv
    fLog_wholeData = fopen(log_wholeDataFileName, 'a+');
    fprintf(fLog_wholeData, '%s\n', subjectWithSession);
    fclose(fLog_wholeData);
end
