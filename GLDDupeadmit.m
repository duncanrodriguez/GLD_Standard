function [DA] = GLDDupeadmit(file)
fid = fopen(file);
i = 0;
disp('Reading GLM...')
Objects = {};
while ~feof(fid)
    newline = fgets(fid);
    if containsRelaxed(newline,'name "')
        string = cellstr(newline);
        stringtemp = regexprep(string,{'name',';',' ','"'},{'','','',''});
        found = 0;
        exit = 0;
        while found == 0 && exit == 0
            newline = fgets(fid);
            if ~containsRelaxed(newline,'}')
                if containsRelaxed(newline,'from "')
                    i = i+1;
                    string = cellstr(newline);
                    string = regexprep(string,{'from',';',' ','"'},{'','','',''});
                    Objects{2}(i) = string;
                    found = 1;
                end
            else
                exit = 1;
            end
        end
        found = 0;
        while found == 0 && exit == 0
            newline = fgets(fid);
            if ~containsRelaxed(newline,'}')
                if containsRelaxed(newline,'to "')
                    string = cellstr(newline);
                    string = regexprep(string,{'to',';',' ','"'},{'','','',''});
                    Objects{3}(i) = string;
                    found = 1;
                end
            else
                exit = 1;
            end
        end
        if exit == 0
            Objects{1}(i) = stringtemp;
        end
    end
end
fclose(fid);
Objects{1} = [Objects{1},Objects{1}];
Objectss{2} = Objects{2};
Objects{2} = [Objects{2},Objects{3}];
Objects{3} = [Objects{3},Objectss{2}];
Objects{1}=transpose(Objects{1});
Objects{2}=transpose(Objects{2});
Objects{3}=transpose(Objects{3});
[Objects{2},I] = sort(Objects{2});
Objects{1} = Objects{1,1}(I,:);
Objects{3} = Objects{1,3}(I,:);

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Find duplicates
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
disp('Finding duplicates...')
k=1;
for i = 1:length(Objects{2})-3
    if strcmp(Objects{2}(i),Objects{2}(i+1))==1
        j=i+1;
        while strcmp(Objects{2}(i),Objects{2}(j))==1
            if strcmp(Objects{3}(i),Objects{3}(j))==1
                DupeAdmit(k,1)=Objects{1}(i);
                DupeAdmit(k,2)=Objects{1}(j);
                k=k+1;
            end
            j=j+1;
        end
    end
end
if exist('DupeAdmit') == 1
    [DA(:,1),I] = unique(DupeAdmit(:,1));
else
    DA = {};
end
end