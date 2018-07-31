function [grouping,group] = traverseIsland(allSort,toSort,fromSorting,grouping,group,object)
%traverseTree Summary of this function goes here
%   Detailed explanation goes here
len = length(toSort{1});
curNode{1} = toSort{2}(object);
curNode{2} = toSort{3}(object);
for node = 1:2 %first connection then second
    for nlist = 2:3 %
        parentlist = binsearch(allSort{nlist},curNode{node});
        if parentlist == 0
            continue
        end
        ii = 1;
        while parentlist(1)+ii<len && strcmp(allSort{nlist}(parentlist(1)+ii),curNode{node})
            parentlist(1+ii) = parentlist(1)+ii;
            ii = ii+1;
        end
        iii = 1;
        while parentlist(1)-iii>0 && strcmp(allSort{nlist}(parentlist(1)-iii),curNode{node})
            parentlist(1+ii+iii-1) = parentlist(1)-iii;
            iii = iii+1;
        end
        if nlist == 2
            parentlist = fromSorting(parentlist);
        end
        self = find(parentlist==object);
        if ~isempty(self)
            for s = length(self):-1:1
                parentlist(self(s)) = [];
            end
        end
        if isempty(parentlist)
            continue
        end
        for j = 1:length(parentlist)
            parent = parentlist(j);
            if grouping(parent) == 0 %parent of current node is not in a group
                if grouping(object) == 0 %current node is not in a group
                    group = group+1;
                    grouping(object) = group;
                    grouping(parent) = group;
                else
                    grouping(parent) = grouping(object);
                end
                [grouping,group] = traverseIsland(allSort,toSort,fromSorting,grouping,group,parent);
            end
        end
    end
end
end