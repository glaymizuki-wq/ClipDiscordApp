using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using ClipDiscordApp.Models;
using ClipDiscordApp.Services;

namespace ClipDiscordApp.Views
{
    public partial class RuleSettingsWindow : Window
    {
        private List<ExtractRule> _rules = new();
        private ExtractRule? _selected;

        public RuleSettingsWindow()
        {
            InitializeComponent();
            CmbType.ItemsSource = Enum.GetValues(typeof(ExtractRuleType)).Cast<ExtractRuleType>();
            LoadRules();
            WireEvents();
            RefreshList();
        }

        private void WireEvents()
        {
            BtnAdd.Click += BtnAdd_Click;
            BtnDuplicate.Click += BtnDuplicate_Click;
            BtnRemove.Click += BtnRemove_Click;
            BtnMoveUp.Click += BtnMoveUp_Click;
            BtnMoveDown.Click += BtnMoveDown_Click;
            BtnValidateSelected.Click += BtnValidateSelected_Click;
            BtnSave.Click += BtnSave_Click;
            BtnApply.Click += BtnApply_Click;
            BtnTestMatch.Click += BtnTestMatch_Click;
            ListRules.SelectionChanged += ListRules_SelectionChanged;
            TxtFilter.TextChanged += (s, e) => RefreshList();

            // ここから編集領域の変更を検知するハンドラを追加
            TxtName.TextChanged += EditorContentChanged;
            TxtPattern.TextChanged += EditorContentChanged;
            TxtOrder.TextChanged += EditorContentChanged;
            TxtPreviewInput.TextChanged += (s, e) => { /* テスト用なら不要 */ };
            ChkEnabled.Checked += EditorContentChanged;
            ChkEnabled.Unchecked += EditorContentChanged;
            CmbType.SelectionChanged += EditorContentChanged;

            BtnApply.IsEnabled = false;
        }
        private void LoadRules()
        {
            _rules = RuleStore.Load();
            if (_rules == null) _rules = new List<ExtractRule>();
        }

        private void RefreshList()
        {
            var filter = TxtFilter.Text?.Trim() ?? "";
            IEnumerable<ExtractRule> view = _rules;
            if (!string.IsNullOrEmpty(filter))
            {
                view = view.Where(r => r.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                    || r.Pattern.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            ListRules.ItemsSource = view.OrderBy(r => r.Order).ToList();
            if (ListRules.Items.Count > 0 && ListRules.SelectedIndex < 0) ListRules.SelectedIndex = 0;
        }

        private void ListRules_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ListRules.SelectedItem is ExtractRule r)
            {
                _selected = r;
                BindToEditor(r);
            }
            else
            {
                _selected = null;
                ClearEditor();
            }
        }

        private void EditorContentChanged(object? sender, RoutedEventArgs e)
        {
            // 編集中に選択がないなら無効化のまま
            if (_selected == null) { BtnApply.IsEnabled = false; return; }

            // 現在エディタの値と選択ルールの実際の値が異なれば有効化
            bool dirty = false;
            if (ChkEnabled.IsChecked != _selected.Enabled) dirty = true;
            if ((ExtractRuleType?)CmbType.SelectedItem != _selected.Type) dirty = true;
            if (TxtName.Text != (_selected.Name ?? "")) dirty = true;
            if (TxtPattern.Text != (_selected.Pattern ?? "")) dirty = true;
            if (!int.TryParse(TxtOrder.Text, out var ord) || ord != _selected.Order) dirty = true;

            BtnApply.IsEnabled = dirty;
        }

        private void BindToEditor(ExtractRule r)
        {
            ChkEnabled.IsChecked = r.Enabled;
            CmbType.SelectedItem = r.Type;
            TxtName.Text = r.Name;
            TxtPattern.Text = r.Pattern;
            TxtOrder.Text = r.Order.ToString();
            BtnApply.IsEnabled = false;
            TxtPreviewResult.Text = "";
        }

        private void ClearEditor()
        {
            ChkEnabled.IsChecked = false;
            CmbType.SelectedIndex = 0;
            TxtName.Text = "";
            TxtPattern.Text = "";
            TxtOrder.Text = "0";
            BtnApply.IsEnabled = false;
            TxtPreviewResult.Text = "";
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
            var r = new ExtractRule { Name = "New Rule", Pattern = "", Type = ExtractRuleType.Regex, Enabled = true, Order = (_rules.Any() ? _rules.Max(x => x.Order) + 1 : 0) };
            _rules.Add(r);
            RefreshList();
            SelectRule(r);
        }

        private void BtnDuplicate_Click(object sender, RoutedEventArgs e)
        {
            if (_selected == null) return;
            var copy = new ExtractRule
            {
                Name = _selected.Name + " Copy",
                Pattern = _selected.Pattern,
                Type = _selected.Type,
                Enabled = _selected.Enabled,
                Order = (_rules.Any() ? _rules.Max(x => x.Order) + 1 : 0)
            };
            _rules.Add(copy);
            RefreshList();
            SelectRule(copy);
        }

        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            if (_selected == null) return;
            if (System.Windows.MessageBox.Show($"Delete rule '{_selected.Name}'?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            _rules.Remove(_selected);
            _selected = null;
            RefreshList();
            ClearEditor();
        }

        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            if (_selected == null) return;
            var idx = _rules.OrderBy(x => x.Order).ToList().FindIndex(x => x.Id == _selected.Id);
            if (idx <= 0) return;
            var ordered = _rules.OrderBy(x => x.Order).ToList();
            var tmp = ordered[idx - 1].Order;
            ordered[idx - 1].Order = ordered[idx].Order;
            ordered[idx].Order = tmp;
            foreach (var r in ordered) _rules.Find(x => x.Id == r.Id)!.Order = r.Order;
            RefreshList();
            SelectRule(_selected);
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            if (_selected == null) return;
            var ordered = _rules.OrderBy(x => x.Order).ToList();
            var idx = ordered.FindIndex(x => x.Id == _selected.Id);
            if (idx < 0 || idx >= ordered.Count - 1) return;
            var tmp = ordered[idx + 1].Order;
            ordered[idx + 1].Order = ordered[idx].Order;
            ordered[idx].Order = tmp;
            foreach (var r in ordered) _rules.Find(x => x.Id == r.Id)!.Order = r.Order;
            RefreshList();
            SelectRule(_selected);
        }

        private void BtnValidateSelected_Click(object sender, RoutedEventArgs e)
        {
            if (_selected == null) return;
            if (_selected.Type != ExtractRuleType.Regex)
            {
                System.Windows.MessageBox.Show("Selected rule is not Regex type.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            ValidateRegex(_selected.Pattern, showMessage: true);
        }

        private bool ValidateRegex(string pattern, bool showMessage)
        {
            try
            {
                _ = new Regex(pattern);
                if (showMessage) System.Windows.MessageBox.Show("Regex is valid.", "OK", MessageBoxButton.OK, MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                if (showMessage) System.Windows.MessageBox.Show($"Invalid regex: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (BtnApply.IsEnabled) ApplyEditorToSelected();

            foreach (var r in _rules.Where(x => x.Enabled && x.Type == ExtractRuleType.Regex))
            {
                if (!ValidateRegex(r.Pattern, showMessage: true))
                {
                    return;
                }
            }

            var ordered = _rules.OrderBy(x => x.Order).ToList();
            for (int i = 0; i < ordered.Count; i++) ordered[i].Order = i;

            RuleStore.Save(ordered);
            System.Windows.MessageBox.Show("Rules saved", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
            LoadRules();
            RefreshList();
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            ApplyEditorToSelected();
        }

        private void ApplyEditorToSelected()
        {
            if (_selected == null) return;
            _selected.Enabled = ChkEnabled.IsChecked == true;
            _selected.Type = (ExtractRuleType)CmbType.SelectedItem!;
            _selected.Name = TxtName.Text.Trim();
            _selected.Pattern = TxtPattern.Text;
            if (int.TryParse(TxtOrder.Text, out var ord)) _selected.Order = ord;
            BtnApply.IsEnabled = false;
            RefreshList();
            SelectRule(_selected);
        }

        private void BtnTestMatch_Click(object sender, RoutedEventArgs e)
        {
            if (_selected == null)
            {
                TxtPreviewResult.Text = "No rule selected";
                return;
            }

            var input = TxtPreviewInput.Text ?? "";
            if (string.IsNullOrEmpty(input))
            {
                TxtPreviewResult.Text = "Paste OCR text to test";
                return;
            }

            if (_selected.Type == ExtractRuleType.Keyword)
            {
                var ok = input.IndexOf(_selected.Pattern ?? "", StringComparison.OrdinalIgnoreCase) >= 0;
                TxtPreviewResult.Text = ok ? "MATCH" : "NO MATCH";
            }
            else
            {
                try
                {
                    var rx = new Regex(_selected.Pattern ?? "", RegexOptions.IgnoreCase);
                    var ms = rx.Matches(input);
                    if (ms.Count == 0)
                    {
                        TxtPreviewResult.Text = "NO MATCH";
                    }
                    else
                    {
                        var list = string.Join("; ", ms.Cast<System.Text.RegularExpressions.Match>().Select(m => m.Value));
                        TxtPreviewResult.Text = $"MATCH: {list}";
                    }
                }
                catch (Exception ex)
                {
                    TxtPreviewResult.Text = $"Invalid regex: {ex.Message}";
                }
            }
        }

        private void SelectRule(ExtractRule r)
        {
            ListRules.SelectedItem = r;
            ListRules.ScrollIntoView(r);
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}