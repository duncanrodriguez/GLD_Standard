% HOW TO USE THIS FILE
%   1) Run the file and you will be prompted to use prior settings or not
%   2) This allows you to choose which GLM file you wish to graph or use
%   the most recent file
%   3) Prompts for the type of graph and how nodes will be displayed will
%   follow, though consider the options' pros/cons below
%
%   Tree: This method is very fast but can only handle a pure radial tree;
%   objects can only have a single parent. This should only be used to
%   draw systems which are known to be radial or to gain a mere sense of
%   the structure.
%
%   Biograph: This method can handle a system of n objects with up to
%   (n^2-n) connections, so multiple parents are allowed, but the graph is
%   very slow to draw (also related to n^2). Thus, this is required to
%   understand the structure of meshed networks.
%
%   Written by Duncan Rodriguez for CWRU

clear all;
%
dirFile = fopen('TreeDir.txt','r');
initDir1 = fgetl(dirFile);
GLM = fgetl(dirFile);
graph = fgetl(dirFile);
val = fgetl(dirFile);
bgLayout = fgetl(dirFile);
fclose(dirFile);
dir = char(initDir1);
usePow = 'Yes';
keepCurrentSettings = questdlg('Use previous file and settings?',...
    'Confirmation',...
    'Yes','No','File only','Yes');
switch keepCurrentSettings
    case 'Yes'
    case 'No'
    case 'File only'
end
if strcmp(keepCurrentSettings,'No')
    %Select the Export file to use and output directory for Jason Fuller files
    % Directory for input files (CSVs)
    [GLM,dir,filter] = uigetfile({'*.glm;*.GLM','All GLD files';'*.*','All Files' },'Select GLM file to graph',dir);
    dirFile = fopen('TreeDir.txt','w');
    fprintf(dirFile,'%s',dir);
    fprintf(dirFile,'\r\n');
    fprintf(dirFile,'%s',GLM);
    fprintf(dirFile,'\r\n');
    fclose(dirFile);
end
if strcmp(keepCurrentSettings,'File only') ||strcmp(keepCurrentSettings,'No')
    dirFile = fopen('TreeDir.txt','w');
    fprintf(dirFile,'%s',dir);
    fprintf(dirFile,'\r\n');
    fprintf(dirFile,'%s',GLM);
    fprintf(dirFile,'\r\n');
    graph = questdlg('Do you want to create a Tree or Biograph?',...
        'Confirmation',...
        'Tree','Biograph','Tree');
    switch keepCurrentSettings
        case 'Tree'
        case 'Biograph'
    end
    fprintf(dirFile,'%s',graph);
    fprintf(dirFile,'\r\n');
    val = questdlg('Do you want the nodes to be Numbered or Named?',...
        'Confirmation',...
        'Number','Name','Number');
    switch keepCurrentSettings
        case 'Number'
        case 'Name'
    end
    fprintf(dirFile,'%s',val);
    fprintf(dirFile,'\r\n');
    if strcmp(graph, 'Biograph')
        bgLayout = questdlg('Do you want to a radial or hierarchical?',...
            'Confirmation',...
            'radial','hierarchical','radial');
        switch keepCurrentSettings
            case 'radial'
            case 'hierarchical'
        end
    end
    fprintf(dirFile,'%s',bgLayout);
    fprintf(dirFile,'\r\n');
    fclose(dirFile);
end
dirr = strcat(dir,GLM);
tic();
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fid = fopen(dirr);
i=0;
disp('Reading GLM...')
while ~feof(fid)
    buffer = '';
    curline = fgets(fid);
    if containsRelaxed(curline,'{')
        buffer = curline;
        curline = fgets(fid);
        buffering = 1;
        while buffering
            if containsRelaxed(curline,'}')
                buffering = 0;
            end
            buffer = [buffer,curline];
            curline = fgets(fid);
        end
        if containsRelaxed(buffer,'from "')
            i = i+1;
            nft = extractBetween(buffer,'"','"');
            Objects{1}(i,1) = nft(1);
            Objects{2}(i,1) = nft(2);
            Objects{3}(i,1) = nft(3);
        end
    end
end
fclose(fid);
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if strcmp(graph,'Tree')
    disp('Building Node List...')
    [Objects{3},I] = sort(Objects{3});
    Objects{1} = Objects{1}(I,:);
    Objects{2} = Objects{2}(I,:);
    len = length(Objects{1});
    
    disp('Building Tree...')
    for i = 1:len
        curNode = Objects{2}(i);
        Objects{4}(i) = binsearch(Objects{3},curNode{1});
    end
    Objects{4} = [Objects{4}];
    disp('Drawing Tree...')
    treeplot(Objects{4},'o','cyan');
    [x,y] = treelayout(Objects{4});
    for i=1:length(x)
        if strcmp(val,'Number')~=1
            txt = Objects{3}(i);
            text(x(i),y(i),txt{1})
        else
            text(x(i),y(i),num2str(i))
        end
    end
end
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if strcmp(graph,'Biograph')
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
    file = ls(dir);
    numFiles = numrows(file);
    k = 3;
    name = '';
    for i = k:1:numFiles
        currfile = strtrim(file(i,:));
        if contains(currfile,'.csv')
            j = k-2;
            file2find = [dir,'\',currfile];
            P = csvread(file2find,10,1);
            currfile = regexprep({currfile},{'.csv','slash'},{'','/'});
            name{j,1} = currfile{1};
            RMSre{j,1} = rms(P(:,1));
            RMSim{j,1} = rms(P(:,2));
            RMSP{j,1} = sign(P(1,1))*sqrt(RMSre{j,1}^2+RMSim{j,1}^2);
            k = k+1;
        end
    end
    [ObjectList,IAob] = sort(Objects{1});
    NodeList = [Objects{2}(IAob,:);Objects{3}(IAob,:)];
    EdonList = [Objects{3}(IAob,:);Objects{2}(IAob,:)];
    Power = zeros(2*len,1);
    for i = 1:length(name)
        nameObject = name(i);
        id = binsearch(ObjectList,nameObject);
        if id
            Power(id) = RMSP{i};
            if RMSP{i}<0
                NodeList{id} = EdonList{id};
                NodeList{id+len} = EdonList{id+len};
            end
        end
    end
    [NodeList,IA] = sort(NodeList);
    ObjectList = [ObjectList;ObjectList];
    ObjectList = ObjectList(IA,:);
    Power = Power(IA,:);
    for i = length(NodeList):-1:2
        if strcmp(NodeList(i),NodeList(i-1))
            while i>1 && strcmp(NodeList(i),NodeList(i-1))
                Power(i-1) = Power(i-1) + Power(i);
                ObjectList(i) = [];
                NodeList(i) = [];
                Power(i) = [];
                i = i-1;
            end
        end
    end
    groupings = [grouping;grouping];
    [NodeList,IA,~] = unique(NodeList);
    ObjectList = ObjectList(IA,:);
    Power = Power(IA,:);
    groupings = num2str(groupings(IA,:));
    lenList = length(NodeList);
    for i = 1:lenList
        groupingsname{i} = [groupings(i),':',num2str(i)];
    end
    disp('Building Connection Matrix...')
    C = zeros(lenList);
    for i = 1:len
        fromNode = Objects{2}(i);
        toNode = Objects{3}(i);
        a = binsearch(NodeList,fromNode{1});
        b = binsearch(NodeList,toNode{1});
        C(a,b) = 1;
    end
    disp('Building Biograph...this may take time')
    if strcmp(val,'Name')
        bg1 = biograph(C,NodeList);
    else
        bg1 = biograph(C,groupingsname);
        set(bg1, 'ShowTextInNodes', 'label')
    end
    if strcmp(bgLayout,'radial')
        set(bg1, 'LayoutType', 'radial')
    end
    if strcmp(usePow,'Yes')
        Power = abs(Power);
        threshpow = 0.1;
        maxpow = floor(max(Power)/1000);
        minpow = floor(min(Power(Power>threshpow)/1000));
        cm = jet(maxpow-minpow);
        for i = 1:lenList
            pow = floor(Power(i)/1000)-minpow;
            if pow <threshpow/1000
                color = [1 1 1];
            else
                color = cm(pow,:);
            end
            set(bg1.nodes(i), 'Color', color)
            set(bg1.nodes(i), 'LineColor', color)
        end
    end
    disp('Drawing Biograph... this will take a long time')
    view(bg1);
end
disp('Done!')
disp(['Elapsed time: ' num2str(toc())]);