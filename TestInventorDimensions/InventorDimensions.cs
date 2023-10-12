using Inventor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace TestInventorDimensions
{
   public class InventorDimensions
   {
      private Application _invApp;
      private DrawingDocument _drwDoc;
      private TransientGeometry _trg;
      
      public InventorDimensions()
      {
         _invApp = Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
         bool ready = _invApp.Ready;
         _trg = _invApp.TransientGeometry;
      }

      public void DetectDimensions()
      {
         _drwDoc = _invApp.ActiveDocument as DrawingDocument ?? throw new Exception("No Inventor IDW found");
         Sheet sheet = _drwDoc.ActiveSheet;
         foreach (DrawingDimension dim in sheet.DrawingDimensions)
         {
            if (dim.DimensionLine is LineSegment2d line2d)
            {
               CheckDimension2d(line2d, dim.Text.Origin, out var startPoint, out var endPoint, out var textPoint);
               // TODO: check if textPoint is near startPoint or endPoint or in between those points
               //
            }
            else
            if (dim.DimensionLine is LineSegment line)
               CheckDimension3d(line);
         }
      }
      
      private void CheckDimension2d(LineSegment2d line, Point2d txtOrign, out Point2d rotStartPoint, out Point2d rotEndPoint, out Point2d rotText)
      {
         var matrix = _trg.CreateMatrix2d();
         var center = _trg.CreatePoint2d();
         center.X = line.MidPoint.X;
         center.Y = line.MidPoint.Y;

         var lineVector = line.StartPoint.VectorTo(line.EndPoint);
         var horizon = _trg.CreateVector2d(1, 0);
         
         var angle = lineVector.AngleTo(horizon);
         var degree = RadianToDegree(angle);
         angle = DegreeToRadian(-1 * degree);
         
         matrix.SetToRotation(angle, center);
         
         rotStartPoint = _trg.CreatePoint2d(line.StartPoint.X, line.StartPoint.Y);
         rotEndPoint = _trg.CreatePoint2d(line.EndPoint.X, line.EndPoint.Y);
         rotText = _trg.CreatePoint2d(txtOrign.X, txtOrign.Y);
         
         Console.WriteLine("Origin P1: " + rotStartPoint.X + " " + rotStartPoint.Y);
         Console.WriteLine("Origin P2: " + rotEndPoint.X + " " + rotEndPoint.Y);
         Console.WriteLine("Origin TextPoint: " + txtOrign.X + " " + txtOrign.Y);

         Console.WriteLine("-----------");
         rotStartPoint.TransformBy(matrix);
         Console.WriteLine("Rotated P1: " + rotStartPoint.X + " " + rotStartPoint.Y);
         
         rotEndPoint.TransformBy(matrix);
         Console.WriteLine("Rotated P2: " + rotEndPoint.X + " " + rotEndPoint.Y);
         
         rotText.TransformBy(matrix);
         Console.WriteLine("Rotated TextPoint: " + rotText.X + " " + rotText.Y);
      }

      private void CheckDimension3d(LineSegment line)
      {
         // TODO: implement CheckDimension3d
      }
      
      public static double DegreeToRadian(double angle)
      {
         return Math.PI * angle / 180.0;
      }

      public static double RadianToDegree(double angle)
      {
         return angle * (180.0 / Math.PI);
      }

   }
}