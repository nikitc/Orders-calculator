   M S B A M L   `   `   `  ����  OrdersCalcutator!    OrdersCalcutator.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	     � Title$+  %Калькулятор заказов��5   	   $	��400��6%   $	��525��62   .��5      � -   5         �Name$ 
LayoutGrid��5      .+�5	   
   �� -   5         ��Name$ FilePath��5	       �� TextWrapping$
 Wrap=�6$    �� AcceptsReturn . " �� VerticalScrollBarVisibility$
 Auto=�6M   $	��250��6p   $	��100��6|   $��	46,36,0,0q�6�   $
��Left=�6�   $	��Top=�6�   �� 5
   
   -   5         ��Name$ СhooseFileButton��5
      $	��100��6,   $��20��68   $��Выбрать файл��6D   $
��Left=�6t   $	��Top=�6�   $��
321,36,0,0q�6�   �� 5   
   $'��!Фильтр по датам:  с��6   $
��Left=�6.   $��
45,168,0,0q�6I   $	��Top=�6]   (  "System.Windows.Controls.DatePicker  5   
   -   5        $��	StartDate��5      $
��Left=�6&   $��166,170,0,0q�6A   $	��Top=�6V   $	��130��6n   �� 5   
   $
��по��6   $
��Left=�6   $��307,170,0,0q�68   $	��Top=�6M    J� RenderTransformOrigin$ 1.875,0.654&�6e     5   
   -   5        $��
FinishDate��5      $
��Left=�6'   $��341,170,0,0q�6B   $	��Top=�6W   $	��130��6o   �� 5   
   -   5        $ СalcResultButton��5      $%��Посчитать заказы��6+   $
��Left=�6F   $��187,224,0,0q�6a   $	��Top=�6�   $	��156��6�   $��28��6�   $��14(�6�   �� 5   
   -   5        $��
ResultText��5      $�� ��6&   $
��Left=�6.   $
��Wrap=�6I   $��156,270,0,0q�6]   $	��Top=�6r   $	��231��6�   �� 5   
   $��Файлы:��6   $
��Left=�6!   $	��Top=�6<   $��	45,10,0,0q�6T   5      5      