function [] = GLDrecorder(dir,filename,recorder)
glm = [dir,'\Recorders\'];
try
mkdir(dir,'\Recorders\');
catch
end
fid = fopen([dir,'\',filename,'.glm']);
[a, b, c, d, n] = deal(0);
lines = {};
fuses = {};
reclosers = {};
switches = {};
transformers = {};
m = 0;
while ~feof(fid)
    newline = fgets(fid);
    if containsRelaxed(newline,'{')~=0
        if containsRelaxed(newline,'object overhead_line {')~=0
            n = n+1;
            newline = fgets(fid);
            newline = fgets(fid);
            name = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            newline = fgets(fid);
            nameparent = regexprep(newline,{'from',' ','"',';'},{'','','',''});
            Objects{1}(n) = cellstr(name);
            Objects{2}(n) = cellstr(nameparent);
            %lines{2}(a) = cellstr(nameparent);
        elseif containsRelaxed(newline,'object underground_line {')~=0
            n = n+1;
            newline = fgets(fid);
            newline = fgets(fid);
            name = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            newline = fgets(fid);
            nameparent = regexprep(newline,{'from',' ','"',';'},{'','','',''});
            Objects{1}(n) = cellstr(name);
            Objects{2}(n) = cellstr(nameparent);
            %lines{2}(a) = cellstr(nameparent);
        elseif containsRelaxed(newline,'object fuse {')~=0
            n = n+1;
            newline = fgets(fid);
            name = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            newline = fgets(fid);
            newline = fgets(fid);
            nameparent = regexprep(newline,{'from',' ','"',';'},{'','','',''});
            Objects{1}(n) = cellstr(name);
            Objects{2}(n) = cellstr(nameparent);
            %fuses{2}(b) = cellstr(nameparent);
        elseif containsRelaxed(newline,'object recloser {')~=0
            n = n+1;
            newline = fgets(fid);
            name = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            newline = fgets(fid);
            newline = fgets(fid);
            nameparent = regexprep(newline,{'from',' ','"',';'},{'','','',''});
            Objects{1}(n) = cellstr(name);
            Objects{2}(n) = cellstr(nameparent);
            %reclosers{2}(c) = cellstr(nameparent);
        elseif containsRelaxed(newline,'object switch {')~=0
            n = n+1;
            newline = fgets(fid);
            name = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            newline = fgets(fid);
            nameparent = regexprep(newline,{'from',' ','"',';'},{'','','',''});
            Objects{1}(n) = cellstr(name);
            Objects{2}(n) = cellstr(nameparent);
            %switches{2}(d) = cellstr(nameparent);
        elseif containsRelaxed(newline,'object transformer {')~=0
            n = n+1;
            newline = fgets(fid);
            newline = fgets(fid);
            name = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            newline = fgets(fid);
            nameparent = regexprep(newline,{'from',' ','"',';'},{'','','',''});
            Objects{1}(n) = cellstr(name);
            Objects{2}(n) = cellstr(nameparent);
        elseif containsRelaxed(newline,'object regulator {')~=0
            n = n+1;
            newline = fgets(fid);
            name = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            newline = fgets(fid);
            newline = fgets(fid);
            nameparent = regexprep(newline,{'from',' ','"',';'},{'','','',''});
            Objects{1}(n) = cellstr(name);
            Objects{2}(n) = cellstr(nameparent);
        elseif containsRelaxed(newline,'object series_reactor {')~=0
            n = n+1;
            newline = fgets(fid);
            name = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            newline = fgets(fid);
            newline = fgets(fid);
            nameparent = regexprep(newline,{'from',' ','"',';'},{'','','',''});
            Objects{1}(n) = cellstr(name);
            Objects{2}(n) = cellstr(nameparent);
            %transformers{2}(n) = cellstr(nameparent);
        elseif containsRelaxed(newline,'object node ')~=0
            m = m+1;
            newline = fgets(fid);
            newline = fgets(fid);
            if containsRelaxed(newline,'bustype SWING')~=0
                newline = fgets(fid);
            end
            newline = regexprep(newline,{'name',' ','"',';'},{'','','',''});
            Nodes{1}(m) = cellstr(newline);
        end
    end
end
justnodes = 0;
if justnodes == 0
    j = 1;
    fidglm = fopen([glm,recorder,num2str(j),'.glm'],'wt');
    for i = 1:length(Objects{1})
        if ~mod(i,500)
            j = j+1;
            fidglm = fopen([glm,recorder,num2str(j),'.glm'],'wt');
        end
        obj = Objects{1}(i);
        parent = Objects{2}(i);
        objfile = regexprep(obj,{'/'},{'slash'});
        fprintf(fidglm, 'object multi_recorder {\n');
        fprintf(fidglm,'	file %s.csv;\n', objfile{1});
        fprintf(fidglm,'	interval 900;\n');
        fprintf(fidglm,'	limit 100;\n');
        fprintf(fidglm,'	property %s:power_in.real,%s:power_in.imag,%s:voltage_A.real,%s:voltage_A.imag;\n', obj{1},obj{1},parent{1},parent{1});
        fprintf(fidglm,'}\n');
    end
else
    fidglm = fopen([glm,'recorderallnodes.glm'],'wt');
    obj = Nodes{1};
    for j = 1:length(obj)
        objfile = regexprep(obj(j),{'CNXN1'},{''});
        ob = objfile{1};
        objfile = regexprep(obj(j),{'/'},{'slash'});
        fprintf(fidglm, 'object multi_recorder {\n');
        fprintf(fidglm,'	file %s.csv;\n', objfile{1});
        fprintf(fidglm,'	interval 900;\n');
        fprintf(fidglm,'	limit 100;\n');
        fprintf(fidglm,'	property %s:power_in.real,%s:power_in.imag;\n', ob,ob);
        fprintf(fidglm,'}\n');
    end
end
end