% Matlab打开及编辑ppt案例
% Start date: 2022-02-12
% Author: Sun Guofang

clear
close all
clc
% 导入相关类
import mlreportgen.ppt.*
% Presentation(A,B),其中A表示新建的ppt,B为模板，若没有B，则是新建空白ppt
ppt = Presentation('共聚焦问题_new.pptx','共聚焦问题back.pptx');
% 打开ppt
open(ppt);
% 通过add增加ppt，注意需要有说明，说明为ppt版式中类型，注意中英文
titleSlide = add(ppt,'空白');
% 关闭并保存ppt
close(ppt);
% 打开ppt
rptview(ppt)
