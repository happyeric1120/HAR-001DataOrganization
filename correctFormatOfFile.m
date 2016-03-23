function [success excelFileName] = correctFormatOfFile(fileName)
[pathstr,name,ext] = fileparts(fileName);
logFileName = [pathstr '\log.txt'];
csvFileName = [pathstr '\' name '.csv'];
excelFileName = [pathstr '\' name '.xlsx'];
success = 1;
try
    data = textread(fileName, '%s', 'delimiter', '\n');
    fd = fopen(csvFileName, 'w');
    % Read the lines one by one
    for i = 1:length(data)
        % Separate the data by delimiters (tab and space)
        temp = strsplit(data{i}, {'\t', ' '});

        for j = 1:length(temp)
            fprintf(fd, '%s,', temp{j});
        end
        fprintf(fd, '\n');
    end
    fclose(fd);
    % Now we have to convert the text file to excel format
    [dummy, dummy, raw] = xlsread([fileName '.csv']);
    xlswrite(excelFileName, raw);
catch
    success = 0; % Any error
    excelFileName = 'none';
    ferror = fopen(logFileName, 'a+');
    fprintf(ferror, datestr(datetime('now')));
    fprintf(ferror, ' ---------- ');
    fprintf(ferror, '%s has error of corrected format covertion!\n', name);
    fclose(ferror);
end