   M S B A M L   `   `   `  ����O  KStavljanje na stanje, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null;"clr-namespace:Stavljanje_na_stanjeStavljanje_na_stanje  5   	   %    Stavljanje_na_stanje.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   4local"clr-namespace:Stavljanje_na_stanje     5   	   Z VFontAwesome.WPF, Version=4.7.0.37774, Culture=neutral, PublicKeyToken=0758b07a11a4f466-fa$http://schemas.fontawesome.io/icons/  5   	     � Title$  
MainWindow��5	   	   $��452.439��6   $��579.269��6-   .��5
      � ���#E1EBEB$��0,0,2,2q�6     � ColumnDefinitions 5   
   �� 5      $
��419* �6   �� 5      $
��223* �6   �� 5      $��0* �6   .+�5   
   �� -   5         ��Name$ btnUvuciFakturu��5      $��- Uvucite fakturu -��6*   $
��Left=�6H   $��54��5      $��
115,58,0,0q�6   $	��Top=�6-   $	��357��5      ���	#FF1CA1A9$��21(�60   ���	#FFE1EBEB$
��Bold"�6j   $��2��6|   �� 5   
   -   5        $ btnSacuvajFajl��5      $��- Sacuvajte fajlove -��6)   $
��Left=�6I   $��
115,76,0,0q�5      $	��Top=�5      $��54��5      $	��357��5      ���	#FF1CA1A9���	#FFE1EBEB$��21(�5      $
��Bold"�6   $��2��61   6E   �� 5   
   -   5        $��Restart��5      $	��Top=�5      $��54��5      $	��170��5      ���	#FF1CA1A9���	#FFE1EBEB$��21(�5"      $
��Bold"�6   $
��Left=�5#      $��115,314,0,0q�6,   6A   �� 5$   
   -   5        $��2��5$      $��Zatvori��6%   $��54��5%      $	��170��5&      ���	#FFD35F5F���	#FFE1EBEB$��21(�5)      $
��Bold"�5*      $
��Left=�5+      $��302,314,0,0q�6,   $	��Top=�6A   6Y   	� 5,   
   -   5         	�Name$ rdbBezOpisa��5,      $��	Bez Opisa��6+   $
��Left=�6?   $��115,198,0,0q�6Z   $	��Top=�6o   $��33��6�   $	��109��6�    d� 	IsChecked$
 TrueY�6�   $��18(�6�   ���	#FF1CA1A9$
��Bold"�6�   $y�Hidden=�6�   	� 5-   
   -   5        $ rdbOpis��5-      $��	Sa Opisom��6'   $
��Left=�6;   $��329,198,0,0q�6V   $	��Top=�6k   $��2��6�   $	��115��6�   $��18(�6�   $
��Bold"�6�   ���	#FF1CA1A9$y�Hidden=�6�   "  FontAwesome.WPF.ImageAwesome  5.   
   -   5          Name$ loadingAnimation��5.         Icon$ Spinner=�64      Spin . $��193,55,182,192q�6O    J� Opacity$	 0.6Q�6g   	   SpinDuration$	 2Q�6u   $��2��6�   �� 5/   
   -   5        
 ��Name$
 	lblObrada��5/      $1��+Obrada je u toku. Molimo sacekajte minut...��6#   $
��Left=�6Y   $��102,236,0,0q�6t   $	��Top=�6�   $��35��6�   $	��420��6�   $��18(�6�   $��2��6�   ���	#FF1CA1A9$
��Bold"�6�   51      52      