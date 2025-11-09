// File: Core/Help.UiProgress.cs
using System;
using System.Windows.Forms;

namespace BetaTestSupp.Core
{
    /// Riuso: gestisce bottone disabilitato + ProgressBar visibile con avanzamento.
    public static class HelpUiProgress
    {
        /// Esegue un'azione con avanzamento UI.
        /// - setup: imposta Maximum (>=1) e progress a 0
        /// - step(): chiamala ad ogni avanzamento (incrementa di 1 e DoEvents)
        public static void Run(Button triggerButton, ProgressBar bar, int maximum, Action<Action> body)
        {
            if (maximum <= 0) maximum = 1;

            triggerButton.Enabled = false;
            bar.Visible = true;
            bar.Style = ProgressBarStyle.Blocks;
            bar.Maximum = maximum;
            bar.Value = 0;

            try
            {
                void step()
                {
                    if (bar.Value < bar.Maximum) bar.Value++;
                    Application.DoEvents();
                }

                body(step);
            }
            finally
            {
                bar.Visible = false;
                triggerButton.Enabled = true;
            }
        }
    }
}
