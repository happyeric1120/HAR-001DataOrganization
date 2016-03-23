clear off
fileName = '\\R04BRXNAS20.v03.med.va.gov\VHABRXWuY$\IRM\Desktop\15-05 S6 MEP (KG).xlsx';
sheetName = 'proc';
[data txt] = xlsread(fileName, sheetName);
header = txt(1,:);
% baseline = -1, 0, 15, 30 mins after STDP ==> Time is column 2
timeStamp = unique(data(:,2));
timeStamp(isnan(timeStamp)) = [];
% channel = 0, 1, 2, 3 ==> Channel is column 6
channel = unique(data(:,6));
channel(isnan(channel)) = [];
% NormIntensity --> column 5, AMPWAVE1 --> column 16

% Group data according to timeStamp
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

iteration = 1;
for i = 1:length(groupDataBasedOnTime)
    tempData = groupDataBasedOnTime{i};
    
    % deal with each channel
    for j = 1:length(channel)
        % channel --> column 6
        channelData = tempData(find(tempData(:,6)== channel(j)),:);
        normIntensity = unique(channelData(:,5));
        
        for k = 1:length(normIntensity)
            % normIntensity --> column 5
            sameIntensityData = channelData(find(channelData(:,5) == normIntensity(k)),:);
            % Performing average for all values (column 16)
            dimensionOfData = size(sameIntensityData);
            % check the dimension of the data. If only one row, don't do
            % average
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
saveHeader = header([2, 5:end]);

xlswrite(fileName, saveHeader, 'average', 'A1');
xlswrite(fileName, resultData, 'average', 'A2');

