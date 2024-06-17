using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ParameterProgramLoader
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("Enter part number:");
			var partName = Console.ReadLine();

			var partFilePath = $@"C:\Temp\parts\{partName}.txt";

			var partParameters = File.ReadAllLines(partFilePath)
				.Select(p => p.Split(','))
				.ToDictionary(p => p[0], p => p[1]);

			var pcdmisType = Type.GetTypeFromProgID("PCDLRN.Application");
			var pcdmis = Activator.CreateInstance(pcdmisType) as PCDLRN.Application;
			var program = pcdmis.ActivePartProgram;
			var commands = program.Commands;

			for (int i = 1; i <= commands.Count; i++)
			{
				var command = commands.Item(i);
				if (command.IsFlowControl)
				{
					var assignmentCommand = command.FlowControlCommand;
					var assignmentName = assignmentCommand.GetLeftSideOfExpression();
					var parameterExists = partParameters.TryGetValue(assignmentName, out string assignmentValue);
					if (parameterExists)
					{
						assignmentCommand.SetRightSideOfAssignment(assignmentValue);
					}

				}
			}

		}
	}
}
