// See https://aka.ms/new-console-template for more information

using System;



Console.WriteLine("Select one of the options:");

Console.WriteLine("   1 Hello World!");
Console.WriteLine("   2 Lewis Chess Piece");
Console.WriteLine("   3 Both");
Console.WriteLine("   e Exit");

var key = Console.ReadKey();
Console.WriteLine("");
if(key.KeyChar is '1' or '3')
    Delaney.Sample.Example1();

if (key.KeyChar is '2' or '3')
    Delaney.Sample.Example2();

if (key.KeyChar is '1' or '2' or '3')
{
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("The documents can be found here:");
    var name = AppDomain.CurrentDomain.BaseDirectory;
    Console.WriteLine(name);
    Console.ResetColor();
}