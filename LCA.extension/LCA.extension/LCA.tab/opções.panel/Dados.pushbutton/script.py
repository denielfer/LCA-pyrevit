# -*- coding: utf-8 -*-
from pyrevit import revit
from System.Windows import Window, Thickness, FontWeights, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Controls import Label, StackPanel, Grid, ColumnDefinition, RowDefinition, ScrollViewer, ScrollBarVisibility, TextBox, Button, Orientation
from System.Windows import HorizontalAlignment, VerticalAlignment, TextAlignment
from data import dataPerKg, salvar_dataPerKg
import re

class LCAWindow(Window):
    def __init__(self):
        self.Title = "LCA Data Input"
        self.MaxWidth = 1280
        self.MaxHeight = 720
        self.Margin = Thickness(10)
        self.headers = ["A1-A3","A4","A5","C1","C2","C3","C4","D"]
        self.__render_()

    def __render_(self):
        self.textboxes = {}

        # Create grid
        grid = Grid()
        

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
        self.label_textButton = tb

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
        self.__make_itens(stack)
    
    def __make_itens(self, stack):
        def on_Delete(code):
            for metric in dataPerKg:
                if (code in dataPerKg[metric]['data']):
                    del dataPerKg[metric]['data'][code]
            self.__render_()
                    
        def on_Edit(code):
            for n,metric in enumerate(dataPerKg, start=1):
                if (code in dataPerKg[metric]['data']):
                    for m,stage in enumerate(self.headers, start=1):
                        self.textboxes[self.__get_textBox_key(n, m)].Text = str(dataPerKg[metric]['data'][code][stage])
                    self.label_textButton.Text = code
        
        desc_added = []
        for metric in dataPerKg:
            for desc in dataPerKg[metric]['data']:
                desc_added.append(desc)
        desc_set = set(desc_added)
        
        for metric in desc_set:
            stack.Children.Add(self.__make_present_item(
                metric,
                on_Delete,
                on_Edit
            ))

    def __make_present_item(self, text, on_delet, on_edit):
        grid = Grid()
        grid.Margin = Thickness(0,10,0,0)

        col_text = ColumnDefinition()
        col_text.Width = GridLength(4.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col_text)

        col_buttons = ColumnDefinition()
        col_buttons.Width = GridLength(1.0, GridUnitType.Star)
        grid.ColumnDefinitions.Add(col_buttons)

        tb = Label()
        tb.Content = text
        tb.Margin = Thickness(5)
        tb.HorizontalAlignment = HorizontalAlignment.Left
        tb.VerticalAlignment = VerticalAlignment.Center
        tb.FontWeight = FontWeights.Bold
        
        Grid.SetColumn(tb, 0)
        grid.Children.Add(tb)

        btn_stack = StackPanel()
        btn_stack.Orientation = Orientation.Horizontal
        btn_stack.HorizontalAlignment = HorizontalAlignment.Right
        Grid.SetColumn(btn_stack, 1)

        btn_edit = Button()
        btn_edit.Content = "Editar"
        btn_edit.Margin = Thickness(2)
        btn_edit.Click += lambda s, e: on_edit(text)

        btn_delete = Button()
        btn_delete.Content = "Deletar"
        btn_delete.Margin = Thickness(2)
        btn_delete.Click += lambda s, e: on_delet(text)

        btn_stack.Children.Add(btn_edit)
        btn_stack.Children.Add(btn_delete)

        grid.Children.Add(btn_stack)

        return grid

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
            
            def on_focus(sender, args):
                if label.Text == "0":
                    label.Text = ""

            def on_lost_focus(sender, args):
                if label.Text == "":
                    label.Text = "0"

            label.GotFocus += on_focus
            label.LostFocus += on_lost_focus

            
            label.PreviewTextInput += only_numbers

        Grid.SetRow(label, row)
        Grid.SetColumn(label, col)
        return label

    def __get_textBox_key(self, row, col):
        return '{} - {}'.format(row, col)

    def on_generate(self, sender, args):
        try:
            for n,metric in enumerate(dataPerKg, start=1):
                dataPerKg[metric]['data'][self.label_textButton.Text] = {
                    "A1-A3": float(self.textboxes[self.__get_textBox_key(n, 1)].Text),
                    "A4": float(self.textboxes[self.__get_textBox_key(n, 2)].Text),
                    "A5": float(self.textboxes[self.__get_textBox_key(n, 3)].Text),
                    "C1": float(self.textboxes[self.__get_textBox_key(n, 4)].Text),
                    "C2": float(self.textboxes[self.__get_textBox_key(n, 5)].Text),
                    "C3": float(self.textboxes[self.__get_textBox_key(n, 6)].Text),
                    "C4": float(self.textboxes[self.__get_textBox_key(n, 7)].Text),
                    "D": float(self.textboxes[self.__get_textBox_key(n, 8)].Text),
                }
            salvar_dataPerKg()
            self.__render_()
        except Exception as e:
            print(e)


# Run window
LCAWindow().ShowDialog()
