function [found] = containsRelaxed(str,pattern)
if ~iscell(pattern)
    found = strfind(str,pattern);
else
    found = cellfun(@(s) ~isempty(strfind(str, s)), pattern);
end
if isempty(found)
    found = 0;
else
    found = 1;
end
end