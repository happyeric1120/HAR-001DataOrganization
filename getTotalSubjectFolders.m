function [totalSubjectFolders] = getTotalSubjectFolders(pathName)
foldersInPath = dir(pathName);
iteration = 1;
for i = 1:length(foldersInPath)
    if foldersInPath(i).isdir == 1 & length(foldersInPath(i).name) > 2
        totalSubjectFolders{iteration} = foldersInPath(i).name;
        iteration = iteration + 1;
    end
end