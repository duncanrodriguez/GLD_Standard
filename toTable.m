function T = toTable(data,titleCol)
matSize = size(data);
col = min(matSize(2),length(data));
try
    len = length(data{1});
    Table = cell(1,col);
    for i = 1:col
        if col==1
            if ~ischar(data{1})
                Table = num2str(data);
            else
                Table = data;
            end
        else
            if ~isempty(data{i})&&~iscell(data{i}(1))
                Table(1:len,i) = cellstr(num2str(data{i}));
            else
                try
                    Table(1:len,i) = data{i};
                catch
                    
                end
            end
        end
    end
    Table = deleteEmptyLines(data{2}(:),Table);
catch
end
T = table;
for i = 1:col
    T.(titleCol{i}) = Table(:,i);
end
end