# -*- coding: utf-8 -*-
from pyrevit import revit
from System.Windows import Window, Thickness, FontWeights, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Controls import Label, StackPanel, Grid, ColumnDefinition, RowDefinition, ScrollViewer, ScrollBarVisibility, TextBox, Button
from System.Windows import HorizontalAlignment, VerticalAlignment, TextAlignment
from data import dataPerKg
from System.Windows.Input import TextCompositionEventArgs
import re

class LCAWindow(Window):
    def __init__(self):
        self.Title = "LCA Data Input"
        self.MaxWidth = 1280
        self.MaxHeight = 720
        self.Margin = Thickness(10)

        self.textboxes = {}

        # Create grid
        grid = Grid()
        self.headers = ["A1-A3","A4","A5","C1","C2","C3","C4","D"]
        

        colDef = ColumnDefinition()
        colDef.SetValue(ColumnDefinition.WidthProperty, GridLength(8.0, GridUnitType.Star))
        # colDef.width = GridLength(8, GridUnitType.Star)
        grid.ColumnDefinitions.Add(colDef)


        for _ in range(len(self.headers)):
            colDef = ColumnDefinition()
            # colDef.width = GridLength(1, GridUnitType.Star)
            colDef.SetValue(ColumnDefinition.WidthProperty, GridLength(1.0, GridUnitType.Star))
            grid.ColumnDefinitions.Add(colDef)

        current_row = 0
        cur_col = 0

        grid.RowDefinitions.Add(RowDefinition())
        grid.Children.Add(self._make_label("Metrica", current_row, 0, bold=True))
        cur_col = 1
        for fase in self.headers:
            grid.Children.Add(self._make_label(fase, current_row, cur_col, bold=True))
            cur_col += 1
        current_row+=1

        for metric in dataPerKg:
            grid.RowDefinitions.Add(RowDefinition())
            grid.Children.Add(self._make_label(metric, current_row, 0, bold=True))
            cur_col = 1
            for _ in self.headers:
                grid.Children.Add(self._make_textBox('0', current_row, cur_col, bold=True))
                cur_col += 1
            current_row+=1


        scroll_viewer = ScrollViewer()
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.Content = scroll_viewer
        # self.Content = grid
        
        stack = StackPanel()
        stack.Margin = Thickness(10)
        stack.VerticalAlignment = VerticalAlignment.Top
        scroll_viewer.Content = stack


        tb = TextBox()
        tb.Text = "Parametro 'Descrição' do tubo"

        def on_focus(sender, args):
            if tb.Text == "Parametro 'Descrição' do tubo":
                tb.Text = ""

        def on_lost_focus(sender, args):
            if tb.Text == "":
                tb.Text = "Parametro 'Descrição' do tubo"

        tb.GotFocus += on_focus
        tb.LostFocus += on_lost_focus

        tb.Margin = Thickness(5)
        tb.FontWeight = FontWeights.Bold
        tb.HorizontalAlignment = HorizontalAlignment.Stretch
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.TextAlignment = TextAlignment.Center
        stack.Children.Add(tb)

        stack.Children.Add(grid)

        # Add Generate button
        btn = Button()
        btn.Content = "Generate Dictionary"
        btn.HorizontalAlignment = HorizontalAlignment.Center
        btn.VerticalAlignment = VerticalAlignment.Center
        btn.Click += self.on_generate
        Grid.SetRow(btn, 3)
        Grid.SetColumnSpan(btn, len(self.headers))
        stack.Children.Add(btn)


    def _make_label(self, text, row, col, bold=False):
        label = Label()
        label.Content = text
        label.Margin = Thickness(5)
        label.HorizontalAlignment = HorizontalAlignment.Left
        label.VerticalAlignment = VerticalAlignment.Center
        if bold:
            label.FontWeight = FontWeights.Bold
        Grid.SetRow(label, row)
        Grid.SetColumn(label, col)
        return label
    
    def _make_textBox(self, text, row, col, bold=False, onlyNumber=True):
        label = TextBox()
        self.textboxes[self.__get_textBox_key(row, col)] = label
        label.Text = text
        label.Margin = Thickness(5)
        label.HorizontalAlignment = HorizontalAlignment.Stretch
        label.VerticalAlignment = VerticalAlignment.Center
        label.TextAlignment = TextAlignment.Center
        if bold:
            label.FontWeight = FontWeights.Bold
        
        if onlyNumber:
            sci_notation_pattern = re.compile(r'^[+-]?\d*\.?\d*(?:[eE][+-]?\d+)?$')
            def only_numbers(sender, e):
                new_text = label.Text[:label.SelectionStart] + e.Text + label.Text[label.SelectionStart + label.SelectionLength:]
                if not sci_notation_pattern.match(new_text):
                    e.Handled = True
            
            label.PreviewTextInput += only_numbers

        Grid.SetRow(label, row)
        Grid.SetColumn(label, col)
        return label

    def __get_textBox_key(self, row, col):
        return '{} - {}'.format(row, col)

    def on_generate(self, sender, args):

        for n,metric in enumerate(dataPerKg, start=1):
            # print("{}: ".format(metric, ...[ self.textboxes[self.__get_textBox_key(n, m)] for m in range(len(self.headers)) ]))
            print("{}: {} {} {} {} {} {} {} {}".format(metric, *[ self.textboxes[self.__get_textBox_key(n, m+1)].Text for m in range(len(self.headers)) ]))

# Run window
LCAWindow().ShowDialog()
