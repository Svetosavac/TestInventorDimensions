using System;


namespace TestInventorDimensions
{
   internal class Program
   {
      public static void Main(string[] args)
      {
         try
         {
            var invDim = new InventorDimensions();
            invDim.DetectDimensions();
         }
         catch(Exception e)
         {
            Console.WriteLine(e.Message);
         }
         
      }
   }
}