clear all;
clc;

[status,sheets] = xlsfinfo('load_profile.xlsx');
numOfSheets = numel(sheets);
% num of sheets ile kaç tane load var onlar? ö?reniyorum
zz=zeros(24,1);
compansation=zeros(1,numOfSheets);

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
        
        if (q(cc,bb)>0) &&((p(cc,bb)/sqrt(p(cc,bb)^2+q(cc,bb)^2))<0.95);
            comp(cc,bb)=q(cc,bb)-(p(cc,bb)/0.95)*sqrt(1-0.95^2);
            %lagging
        elseif (q(cc,bb)<0)&&((p(cc,bb)/sqrt(p(cc,bb)^2+q(cc,bb)^2))<0.98);
            comp(cc,bb)=q(cc,bb)+(p(cc,bb)/0.98)*sqrt(1-0.98^2);
            %lead
        else
            comp(cc,bb)=0;        
    end   
    end
       
end
    % comp fonksiyonu kompanzasyon degerlerini iceriyor.
    comp1=sort(comp);
    %comp1 büyükten kucuge s?ral?yor
    
    for n=1:numOfSheets
        if abs(comp1(24,n))<abs(comp1(1,n));
            compansation(1,n)=comp1(1,n);
        else abs(comp1(24,n))>abs(comp1(1,n));
            compansation(1,n)=comp1(24,n);
        end 
    end
    
    %burdaki for leading pf leri de i?in içine katmak için bulunuyor.
            
  %companzasyon matrixi tüm loadlar için gereken kompanzasyonu veriyor.
  %bu andan sonra tüm loadlar? kompanze edilmi? haline çevirebiliriz.
 
    for bb=1:numOfSheets
    for cc= 1:24
        compansatedload(cc,bb*2-1)=p(cc,bb);
          if (q(cc,bb)>0) &&((p(cc,bb)/sqrt(p(cc,bb)^2+q(cc,bb)^2))<0.95);
            compansatedload(cc,bb*2)=(p(cc,bb)/0.95)*sqrt(1-0.95^2);
        elseif (q(cc,bb)<0)&&((p(cc,bb)/sqrt(p(cc,bb)^2+q(cc,bb)^2))<0.98);
            compansatedload(cc,bb*2)=-(p(cc,bb)/0.98)*sqrt(1-0.98^2);
        else
            comp(cc,bb)=0;   
             
    end 
       
    end
    end
    %compansatedload bize companse edilmis loadlar?n p ve qsuna sahip olan
    %loadlar? veriyor.
    
    
    
    
    %%Buradan sonra kablo tipi seçme i?ine giricez.
    
    cableareas=[240,180,120,70];
    
    
