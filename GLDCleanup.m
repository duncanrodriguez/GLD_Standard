function [] = GLDCleanup(file)
%GLDCleanup Cleans up GLM files by getting rid of islands.
%   As long as the islands of the GLM file are marked with the top node
%   being a swing bus, those swing buses will be found. They will then be
%   pathed to their leaves to determine if each is a small island fit for
%   deletion. A list of these islands is created and then those islands
%   are fully traversed and deleted.


%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
[Nodelist, groupings,groups] = GLDIslandfinder(file);
mainisland = groups( groups(:,2) == max(groups(:,2)),1);

NodeList = {};
j = 1;
for i = 1:length(groupings)
    if groupings(i) ~= mainisland
        NodeList{j} = Nodelist{i};
        j = j+1;
    end
end
GLDNodedelete(file,NodeList);
%%
DupeList = GLDDupeadmit(file);
GLDNodedelete(file,DupeList);
end