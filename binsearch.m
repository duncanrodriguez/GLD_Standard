function index = binsearch(A,element)
if iscell(element)
    element = element{1};
end
left = 1;
right = length(A);
flag = 0;
elen = length(element);
while left <= right
    mid = ceil((left + right) / 2);
    Amid = A{mid};
    if strcmp(Amid,element)
        index = mid;
        flag = 1;
        break;
    end
    Amidlen = length(Amid);
    n = min(Amidlen,elen);
    if elen == n
        eltrunc = element;
        
    else
        eltrunc = element(1:n);
    end
    if Amidlen == n
        Amidtrunc = Amid;
    else
        Amidtrunc = Amid(1:n);
    end
    midgtel = Amidtrunc > eltrunc;
    elgtmid = Amidtrunc < eltrunc;
    aemax = max(midgtel,elgtmid);
    greatest = find(aemax);
    if ~isempty(greatest)
        greatest = greatest(1);
    else
        midgtel = [true,false];
        if elen == n
            greatest = 1;
        else
            greatest = 2;
        end
    end
    if midgtel(greatest) == 1
        right = mid - 1;
    else
        left = mid + 1;
    end
end
if flag == 0
    index = 0;
end
end