clear all

dataPath = '\\R04BRXNAS20.v03.med.va.gov\VHABRXWuY$\IRM\Desktop\Matlab Codes\Data\';
thresholdFileName = [dataPath 'SubjectThreshold.csv'];
macroFile = [dataPath 'ProcessData.xlsb'];


% Get all subjects folder
totalSubjectFolders = getTotalSubjectFolders(dataPath);

counterNonAnalyzedFile = 1;
nonAnalyzedFiles = {};
for folder = totalSubjectFolders
    singleSubjectFolderPath = [dataPath folder{1}];
    singleSubjectData = dir(singleSubjectFolderPath);
    
    % finding log file to get all analyzed files
    fileAnalyzed = {};
    subjectLogFile = [singleSubjectFolderPath '\log.csv'];
    if exist(subjectLogFile, 'file') == 0
        fLog = fopen(subjectLogFile, 'w');
        fclose(fLog);
    end
    fLog = fopen(subjectLogFile, 'r');
    iteration = 1;
    tline = fgetl(fLog);
    while ischar(tline)
        fileAnalyzed{iteration} = tline;
        iteration = iteration + 1;
        tline = fgetl(fLog);
    end
    fclose(fLog);
    
    
    for i = 1:length(singleSubjectData)
        if singleSubjectData(i).isdir ~= 1 & strcmp(singleSubjectData(i).name, 'log.csv') == 0
            % find if the current file name in the analyzed file cell
            [pathstr name ext] = fileparts(singleSubjectData(i).name);
            if isempty(ext) == 1 & isempty(cell2mat(strfind(fileAnalyzed, singleSubjectData(i).name))) == 1
                nonAnalyzedFiles{counterNonAnalyzedFile} = [dataPath folder{1} '\' singleSubjectData(i).name];
                counterNonAnalyzedFile = counterNonAnalyzedFile + 1;
            end
        end
    end 
end


% subjectFolder = '001-09';
% dataFileName = '15-09 S7 MEP';
% fileName = [dataPath subjectFolder '\' dataFileName];

for fileName = nonAnalyzedFiles
    fileName = fileName{1};
    [pathStr dataFileName extension] = fileparts(fileName);
    
    disp([dataFileName '.... is analyzing!']);

    processedFolder = '\processExcel';

    if exist([dataPath processedFolder], 'dir') == 0
        mkdir(dataPath, processedFolder);
    end


    %%%%% Determine which file is being analyzed --> F, MEP or CES
    tempNameSplit = strsplit(dataFileName);
    fileType = tempNameSplit{end};



    %%%%% Clean up the format of the data file
    disp([dataFileName '.... is cleaning with correct format!']);
    [status, fileName] = correctFormatOfFile(fileName);
    [subjectID session TMSRMT CESRMT] = getThresholdFromFileName(fileName, thresholdFileName);

    if status == 1 & TMSRMT ~= 0
        %%%%% Calling excel application in matlab
        ExcelApp = actxserver('Excel.Application');



        %%%%% The ProcessData workbook has to be opened first, so other workbook can
        %%%%% link the macro
        macroWB = ExcelApp.Workbooks.Open(macroFile);

        %%%%% Open the file
        fileWB = ExcelApp.Workbooks.Open(fileName);
    %     ExcelApp.Visible = 1;
        ExcelApp.Run('ProcessData.xlsb!ProcessData', TMSRMT, CESRMT);

        %%%%% close the excel application and release resources
        ExcelApp.Quit
        ExcelApp.release
        
        disp([dataFileName '.... running macro is done!']);


        sheetName = 'proc';
        [data txt] = xlsread(fileName, sheetName);
        header = txt(1,:);
        %%%%% baseline = -1, 0, 15, 30 mins after STDP ==> Time is column 2
        timeStamp = unique(data(:,2));
        timeStamp(isnan(timeStamp)) = [];
        %%%%% channel = 0, 1, 2, 3 ==> Channel is column 6
        channel = unique(data(:,6));
        channel(isnan(channel)) = [];
        %%%%% NormIntensity --> column 5, AMPWAVE1 --> column 16

        %%%%% Group data according to timeStamp
        groupDataBasedOnTime = {};
        for i = 1:length(timeStamp)
            k = 1;
            tempData = [];
            for j = 1:length(data)
                if (data(j,2) == timeStamp(i))
                    tempData(k,:) = data(j,:);
                    k = k+1;
                end
            end
            groupDataBasedOnTime{i} = tempData;
        end

        switch fileType
            case {'MEP', 'CES'}
                iteration = 1;
                for i = 1:length(groupDataBasedOnTime)
                    tempData = groupDataBasedOnTime{i};

                    %%%%% deal with each channel
                    for j = 1:length(channel)
                        %%%%% channel --> column 6
                        channelData = tempData(find(tempData(:,6)== channel(j)),:);
                        normIntensity = unique(channelData(:,5));

                        for k = 1:length(normIntensity)
                            %%%%% normIntensity --> column 5
                            sameIntensityData = channelData(find(channelData(:,5) == normIntensity(k)),:);
                            %%%%% Performing average for all values (column 16)
                            dimensionOfData = size(sameIntensityData);
                            %%%%% check the dimension of the data. If only one row, don't do
                            %%%%% average
                            if dimensionOfData(1) == 1
                                meanValues = sameIntensityData(:, 7:end);
                            else
                                meanValues = mean(sameIntensityData(:, 7:end));
                            end
                %             meanAMPWAVE1 = mean(sameIntensityData(:, 16));
                            indicateValue = [sameIntensityData(1,2) round(sameIntensityData(1,5)) sameIntensityData(1,6)];
                            processData = cat(2, indicateValue, meanValues);
                            if iteration == 1
                                resultData = processData;
                            else
                                resultData = cat(1, resultData, processData);
                            end
                            iteration = iteration + 1;
                %             fprintf('Time: %d, Channel: %d, normIntensity: %d, avgAMPWAVE1: %.4f\n', sameIntensityData(1,2), sameIntensityData(1,6), sameIntensityData(1, 5), meanValues);
                        end
                    end 
                end

            case 'F'
                iteration = 1;
                for i = 1:length(groupDataBasedOnTime)
                    tempData = groupDataBasedOnTime{i};

                    %%%%% deal with each channel
                    for j = 1:length(channel)
                        %%%%% channel --> column 6
                        channelData = tempData(find(tempData(:,6)== channel(j)),:);
                        normIntensity = unique(channelData(:,5));

                        dimensionOfData = size(channelData);
                        if dimensionOfData(1) == 1
                            meanValues = channelData;
                        else
                            meanValues = mean(channelData);
                        end
                        processData = cat(2, meanValues(2), 0);
                        processData = cat(2, processData, meanValues(6:end));

                        if iteration == 1
                            resultData = processData;
                        else
                            resultData = cat(1, resultData, processData);
                        end
                        iteration = iteration + 1;

                    end
                end
        end
        disp([dataFileName '.... is saving the average!']);
        saveHeader = header([2, 5:end]);

        xlswrite(fileName, saveHeader, 'average', 'A1');
        xlswrite(fileName, resultData, 'average', 'A2');
        
        % write down in the log.csv to make sure the file is completed
        fLog = fopen([pathStr '\log.csv'], 'a+');
        fprintf(fLog, '%s\n', dataFileName);
        fclose(fLog);
    end
end













% %%%%% Initialize the Excel application
% ExcelApp = actxserver('Excel.Application');
% workBook = ExcelApp.Workbooks.Open(fileName);
% ExcelApp.Visible = 1;
% 
% %%%%% Switch to the sheet named 'average'
% workBook.Sheets.Item('average').Activate();
% 
% %%%%% Adding chart
% chart = ExcelApp.ActiveSheet.Shapes.AddChart;
% chart.Name = 'TestChart';
% 
% testChart = ExcelApp.ActiveSheet.ChartObjects('testChart');
% testChart.Activate;
% seriesCount = ExcelApp.ActiveChart.SeriesCollection.Count();
% 
% %%%%% Delete the series that automatically create
% for i = 1:seriesCount
%     series = ExcelApp.ActiveChart.SeriesCollection(1);
%     series.Delete();
% end
% 
% %%%%% Choosing the series you want to create
% newSeries = ExcelApp.ActiveChart.SeriesCollection.NewSeries();
% newSeries.XValues = ['=average!$B$2:$B$7'];
% newSeries.Values = ['=average!$M$2:$M$7'];
% ExcelApp.ActiveChart.ChartType = 'xlXYScatterLinesNoMarkers';
% 
% axes = ExcelApp.ActiveChart.Axes(1);
% axes.HasTitle = 1;
% axes.AxisTitle.Caption = 'NormIntensity';
% axes = ExcelApp.ActiveChart.Axes(2);
% axes.HasTitle = 1;
% axes.AxisTitle.Caption = 'AMPWAVE1';
% 
% % ExcelApp.Quit
% % ExcelApp.release