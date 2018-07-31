function [ row ] = numrows( data )
if ischar(data)
    matSize = size(data);
    if matSize(1) ~= 0
        row = matSize(1);
    else
        row = 0;
    end
elseif isnumeric(data)
    matSize = size(data);
    if matSize(1) ~= 0
        row = matSize(1);
    else
        row = 0;
    end
else
    matSize = size(data);
    if matSize(1) ~= 0
        row = length(data{1});
    else
        row = 0;
    end
end
end

