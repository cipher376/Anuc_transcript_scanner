   M S B A M L   `   `   `  ���� R  NFTranscript, Version=1.0.0.0, Culture=neutral, PublicKeyToken=0da51ff7fbc18ad5)clr-namespace:FTranscriptFTranscript      FTranscript.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     +localclr-namespace:FTranscript       � Title$  ANUC Transcript Processor��$	��647��$
��1008��$	��800��$	��600�� � Icon$ resource/logo200x200.ico��
���   �.��� -    �Name$ mainContainer�����	#FFD2D6F7 � ColumnDefinitions �� $
��154* ��� $	��27* ��� $
��257* ��� $	��29* ��� $
��154* � � RowDefinitions �� $��12.603 ��� $��24.913 ��� $��48 ��� $��2* ��� $
��235* ��� $	��28* �.+�� $��3��$��5�����	#FF0088CC.+��� -    ��Name$ 
mlogo2_png��$��80��$��
mlogo2.png��$
��Fill=�$��80�� J� RenderTransformOrigin$ 
1.175,0.45&�$��24,0,0,0q�� $��1��$��4�����	#FF0088CC$��3�� �� $	��11* ��� $
��313* ��� $	��135 � �� $	��11* ��� $
��278* ��� $	��10* �.+��� $
��Auto=�$��1��$��1��.���� -    ��Name$ 	DropLabel�� J� 	AllowDrop . $��WrapWithOverflow=����	#FFF9F9F9$@��:Please drop transcript files here Or Click to select files��	 �� TextAlignment$	 Center=�
 J� OpacityMask
 �	#FFA88A8A���	#FF535D96$��16(��� $��0q�$��2��$��1��.+��� -    ��Name$ genBtn��$��Generate Transcript�����	#FFF9F9F9���	#FF18478A���	#B7222E70$	��0,3q� � ToolTip$ create transcript�� � Cursor$
 Hand��$��Yu Gothic Medium)�$
��Bold"�$��0,20,0,0q�$��17(��� -   $ 
clrViewBtn��$��
Clear View�����	#FFF9F9F9���	#FF18478A���	#B7222E70$	��0,3q�$ create transcript��$
 Hand��$��Yu Gothic Medium)�$
��Bold"�$��0,10,0,0q�$��17(��� -   $ title��$��2��$
��Wrap=�$��Transcript Processor��
�������$��Right=�$��Center=�$��Charlemagne Std)�$��18(�$��	0,12,10,0q�$��2��$��26��$	��308����Z� �� -    ��Name$ footer��$I��CDeveloped by Alfred (ntiamoah376@yahoo.com)      http://natlink.net��$��5��$��5��
��� ���$��0q����	#FF0088CC$��14(��� -    ��Name$ 	statusBar��$��10q�$��5��
�������$��0q����	#FF0088CC$��14(�$��Right=�$
��Bold"�$��4��$��26��.@��� .���� -	   $��	statuslbl���� $��0��$��4��$	��4,0q�.+��� -
   $ headerHeightLbl��$��Header height(mm): 100�����	#FF0088CC
�������$��17(��� -    ��Name$ headerHeightSld��$��0,0,0,25q�
������� � Minimum$ 10Q� � Maximum$	 200Q� �� TickPlacement$ BottomRight=����	#FF535D96 � Value$	 100Q��� -   $ bodyHeightLbl��$��Body height(mm): 141�����	#FF0088CC
�������$��17(��� -   $ bodyHeightSld��$��0,0,0,25q�
�������$	 100Q�$	 200Q�$ BottomRight=� �� TickFrequency$ 1Q����	#FF535D96$
 Hand��$	 141Q��� -   $ footerHeightLbl��$��Footer offset(mm): 20�����	#FF0088CC
�������$��17(��� -   $ footerHeightSld��$��0,0,0,20q�
�������$ 10Q�$ 30Q�$ BottomRight=����	#FF535D96$ 20Q��� -   $ resetSld��$��Reset Slider�����	#FF0088CC
����������	#B7222E70$	��0,3q�$ create transcript��$
 Hand��$��Yu Gothic Medium)�$
��Bold"�$��0,0,0,45q�$��17(��� -   $ viewBtn��$��View PDF Files��$��1��$��3�����	#FF0088CC
����������	#B7222E70$	��0,3q�$ create transcript��$
 Hand��$��Yu Gothic Medium)�$
��Bold"�$��0,54,0,0q�$��17(��� -   $ 	DeletePdf��$��Delete PDF files�����	#FF0088CC
����������	#B7222E70$	��0,3q�$ create transcript��$
 Hand��$��Yu Gothic Medium)�$
��Bold"�$��0,10,0,0q�$��17(��� $��4��$��4��$	��4,0q�.+��� -   $ transcriptItems��$��Formatted pdf files�����	#FF0088CC
�������$��17(��� $��5��$��4��.���� -    ��Name$ formattedFiles�����	#FF0088CC$	��431��