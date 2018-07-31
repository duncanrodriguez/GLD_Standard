function [NodeList,groupings,groups] = GLDIslandfinder(file)
%GLDIslandfinder Finds the nodes of islands in a GLM file.


%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fid = fopen(file);
i=0;
disp('Reading GLM...')
while ~feof(fid)
    newline = fgets(fid);
    if containsRelaxed(newline,'name "')~=0
        i = i+1;
        string = cellstr(newline);
        string = regexprep(string,{'name',';',' ','"'},{'','','',''});
        Objects{1}(i,1) = string;
        found = 0;
        while found == 0
            newline = fgets(fid);
            if containsRelaxed(newline,'from "')~=0
                string = cellstr(newline);
                string = regexprep(string,{'from',';',' ','"'},{'','','',''});
                Objects{2}(i,1) = string;
                found = 1;
            end
        end
        found = 0;
        while found == 0
            newline = fgets(fid);
            if containsRelaxed(newline,'to "')~=0
                string = cellstr(newline);
                string = regexprep(string,{'to',';',' ','"'},{'','','',''});
                Objects{3}(i,1) = string;
                found = 1;
            end
        end
    end
end
fclose(fid);
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
disp('Building Node List...')
[Objects{3},I] = sort(Objects{3});
Objects{1} = Objects{1}(I,:);
Objects{2} = Objects{2}(I,:);
[Objectss{2},IA2] = sort(Objects{2});
Objectss{3} = Objects{3};
len = length(Objects{1});
grouping = zeros(len,1);%group #
group = 1;
disp('Building Tree and Finding Islands...')
for i = 1:len
    if ~grouping(i)
        [grouping,group] = traverseIsland(Objectss,Objects,IA2,grouping,group,i);
    end
end

groups = unique(grouping);
groups(:,2) = 0;
groupsize = size(groups);
for i = 1:groupsize(1)
    for j = 1:length(grouping)
        if grouping(j) == groups(i,1)
            groups(i,2) = groups(i,2)+1;
        end
    end
end
groupings = [grouping;grouping];
[NodeList,IA,~] = unique([Objects{2};Objects{3}]);
groupings = groupings(IA,1);
if length(groups)>1
    disp(['Important: ',num2str(length(groups)-1),' islands were found, and ',num2str(sum(groupings==0)),' individual disconnected components.']);
end
end