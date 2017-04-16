clear all;
clc;

[status,sheets] = xlsfinfo('load_profile.xlsx');
numOfSheets = numel(sheets);
% num of sheets ile kaç tane load var onlar? ö?reniyorum


num=xlsread('load_profile.xlsx',1);
    %ilk loadi num a yazd?r?yorum
    a=num;
    for k= 1:(numOfSheets-1)


    num=xlsread('load_profile.xlsx',k+1);
    
   a=horzcat(a,num); 
   % a matrixi ile tüm loadlar? tek bir matrixe indiriyorum
    end

dimension=size(a);
% a n?n kaç sat?r ve sutun oldugunu ogreniyorum


for aa=1:numOfSheets
    for m=1:24
        p(m,aa)=a(m,(aa*3-1));
        q(m,aa)=a(m,aa*3);                   
     end
end
% p ile q loadlar?n kw ve kvar degerlerini baska bir matrixte topluyor
for bb=1:numOfSheets
    for cc= 1:24
        
        if (q(cc,bb)>0) &&((p(cc,bb)/sqrt(p(cc,bb)^2+q(cc,bb)^2))<0.95)
            comp(cc,bb)=q(cc,bb)-(p(cc,bb)/0.95)*sqrt(1-0.95^2);
        elseif (q(cc,bb)<0) &&((p(cc,bb)/sqrt(p(cc,bb)^2+q(cc,bb)^2))<0.98)
            comp(cc,bb)=q(cc,bb)-(p(cc,bb)/0.95)*sqrt(1-0.98^2);
        else
            comp(cc,bb)=0;        
    end   
    end
       
end
    % comp fonksiyonu kompanzasyon degerlerini iceriyor.
    

    
