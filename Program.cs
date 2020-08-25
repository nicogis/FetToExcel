//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Studio A&T s.r.l.">
//     Author: nicogis
//     Copyright (c) Studio A&T s.r.l. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------
namespace FetToExcel
{
    using System;
    using System.Windows.Forms;

    static class Program
    {
        /// <summary>
        /// Punto di ingresso principale dell'applicazione.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FetToExcel());
        }
    }
}
