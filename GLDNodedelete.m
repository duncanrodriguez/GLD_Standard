function [] = GLDNodedelete( file,Nodelist)
%Nodes set for deletion should be named as "node_name" rather than the GLD
%allowed node_name without parenthesis. This acts as a flag that this is
%the node you are looking for and not another object which containsRelaxed that
%name.
fid = fopen(file);
filetemp = regexprep(file,{'(?=glm).*','\.'},{'','temp.glm'});
fidtemp = fopen(filetemp,'w');
disp('Deleting...');
filetxt = '';
lineNum = [];
if ~isempty(Nodelist)
    filetxt = fileread(file);
    for i = 1:length(Nodelist)
        node = Nodelist{i};
        lineNum = [lineNum,strfind(filetxt,['"',node,'"'])];
        lineNum = [lineNum,strfind(filetxt,[node,'Config'])];
        lineNum = [lineNum,strfind(filetxt,[node,'CNXN1'])];
    end
    lineNum = unique(lineNum);
    lenlineNum = length(lineNum);
    filetxt2 = '';
    seclen = 2^20;
    lensection = 0;
    lenfiletxtold = 0;
    lenbuffer = 0;
    buffer = '';
    curline = '';
    for i = 1:lenlineNum
        charNum = lineNum(i)-100;
        if not(lensection+lenfiletxtold+lenbuffer>lineNum(i))
            while lensection+lenfiletxtold+lenbuffer<charNum
                filetxt2 = [filetxt2,curline];
                curline = fgets(fid);
                lensection = length(filetxt2);
                if lensection>seclen
                    fprintf(fidtemp,'%s',filetxt2);
                    lenfiletxtold = lenfiletxtold + lensection;
                    lensection = 0;
                    filetxt2 = '';
                end
            end
            while ~containsRelaxed(curline,'{')
                filetxt2 = [filetxt2,curline];
                curline = fgets(fid);
            end
            fprintf(fidtemp,'%s',filetxt2);
            lenfiletxtold = lenfiletxtold + length(filetxt2);
            filetxt2 = '';
            lensection = 0;
            buffering = 1;
            while buffering
                if containsRelaxed(curline,'}')%read the next line to get rid of newlines between deleted objects
                    buffering = 0;
                end
                buffer = [buffer,curline];
                curline = fgets(fid);
            end
            lenbuffer = lenbuffer + length(buffer);
            buffer = '';
        end
    end
    filetxt2 = filetxt(lensection+lenfiletxtold+lenbuffer:length(filetxt));
    fprintf(fidtemp,'%s',filetxt2);
    fclose('all');
    movefile(filetemp,file,'f');
end
fclose('all');