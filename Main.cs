using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI3._3
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Walls));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Doors));
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Windows));

            using (Transaction ts = new Transaction(doc, "Add parameter"))
            {
                ts.Start();
                CreateShareParameter(uiapp.Application, doc, "Длина с запасом", categorySet, BuiltInParameterGroup.PG_GEOMETRY, true);
                ts.Commit();
            }

            //var selectedRef = uidoc.Selection.PickObject(ObjectType.Element, "Выберете элемент");
            //var selectedElement = doc.GetElement(selectedRef);

            IList<Reference> selectedRef = uidoc.Selection.PickObjects(ObjectType.Element, new PipeFilter(), "Выберете элементы труб");
            var pipeList = new List<Pipe>();
            var len = new List<double>();
            double sum = 0;
            string Length = string.Empty;

            foreach (var selectedElement in selectedRef)
            {
                Pipe oPipe = doc.GetElement(selectedElement) as Pipe;
                pipeList.Add(oPipe);
            }

            foreach (var element in selectedRef)
            {
                var selectedElement = doc.GetElement(element);
                using (Transaction ts = new Transaction(doc, "Set parameters"))
                {
                    ts.Start();
                    Parameter commentParameter = selectedElement.LookupParameter("Длина с запасом");
                    Parameter length = selectedElement.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    double v = UnitUtils.ConvertFromInternalUnits(length.AsDouble(), UnitTypeId.Meters) * 1.1 / 304.78;
                    commentParameter.Set(v);
                    ts.Commit();
                }
            }

            foreach (var pipe in pipeList)
            {
                Parameter pipeLength = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                double V = UnitUtils.ConvertFromInternalUnits(pipeLength.AsDouble(), UnitTypeId.Millimeters);
                len.Add(V);
            }

            foreach (var i in len)
            {
                sum += i;
            }
            string SumLength = Math.Round(sum,2).ToString();
            string SumLength10 = Math.Round((sum * 1.1),2).ToString();
            Length += $"Длина: {SumLength}мм{Environment.NewLine}Общее количество элементов труб: {pipeList.Count}{Environment.NewLine}" +
                $"Длина с запасом 10%: {SumLength10}мм";

            TaskDialog.Show("Selection", Length);

            return Result.Succeeded;
        }

        private void CreateShareParameter(Application application,
            Document doc, string parameterName, CategorySet categorySet,
            BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile defFile = application.OpenSharedParameterFile();
            if (defFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }

            Definition definition = defFile.Groups
                .SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builtInParameterGroup);
        }
    }
}