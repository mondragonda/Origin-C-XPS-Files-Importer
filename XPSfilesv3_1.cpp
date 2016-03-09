/*------------------------------------------------------------------------------*
 * File Name:	XPSfiles														*
 * Creation: 	9.8.15															*
 * Purpose: 	Provide a class to handle CASA export files						*
 * 				Cabrera Ortega Jesus Efren										*
 * 				efren.cabrera@uabc.edu.mx										*
 *------------------------------------------------------------------------------*/
 

#include <Origin.h>

test()
{
	string experimentFilePath = GetOpenBox("*.txt");
	
	XPSExperimentImporter xpsExperimentImporter();
	
	xpsExperimentImporter.importExperiment(experimentFilePath);
	
	vector<string> experimentNames;
	experimentNames = xpsExperimentImporter.getExperimentsNames();
	
	for(int i = 0; i < xpsExperimentImporter.getNumberOfExperiments(); i++)
		out_str(experimentNames[i]);
	
	out_int("Number of experiments: ", xpsExperimentImporter.getNumberOfExperiments()); 
}

class XPSExperimentImporter
{
public:
	XPSExperimentImporter();
	void importExperiment(string);
	string getExperimentFileName();
	int getNumberOfSamples();
	int getNumberOfExperiments();
	vector<string> getExperimentsNames();
	
private:
	bool importExperimentFile(string);
	bool isValidPath(string);
	void setExperimentFileName();
	void fixColumns(Worksheet);
	string getNameBeforeChar(string, char);
	void setExperimentsNames();
	string getNameBeforeCompound(string);	
	string getCompound(string);
	string experimentFilePath;
	string experimentFileName;
	void bindExperimentData(WorksheetPage, string);
	bool columnHasExperimentName(Column);
	bool columnNameIncludesCompound(string);
	bool experimentNameIsNotAlreadyAdded(vector<string>, string);
	void addExperimentBEColumnIdentifier();
	void addExperimentBackgroundColumnIdentifier();
	void addExperimentEnvelopeColumnIdentifier();
	void addExperimentColumnIdentifier();
    void separateExperiments();
    void bindLayer(string, Worksheet, string);
	int numberOfSamples;
	vector<string> experimentsNames;
	Worksheet experimentWorksheet;
};

XPSExperimentImporter::XPSExperimentImporter()
{
}

void XPSExperimentImporter::importExperiment(string _experimentFilePath)
{
   if (isValidPath(_experimentFilePath))
	{
		experimentFilePath = _experimentFilePath;
		importExperimentFile(experimentFilePath);
	}
}

bool XPSExperimentImporter::isValidPath(string path)
{
	if (path.IsFile() && path.Match("*.txt"))
		return true;
	else
		return false;
}

bool XPSExperimentImporter::importExperimentFile(string path)
{
	ASCIMP	asciiImporter;
	WorksheetPage wksPg("Book1");

	if (wksPg)
		wksPg.Destroy();

	if (path.IsEmpty())
		return false;

	if(AscImpReadFileStruct(path, &asciiImporter) == 0)
	{
		Worksheet wks;
		wks.Create("Experiments data",CREATE_TEMP);

		int importOperationResult = wks.ImportASCII(path, asciiImporter);
		
		if(importOperationResult == 0)
		{
			experimentWorksheet = wks;
			fixColumns(wks);
			setExperimentsNames();
			addExperimentBEColumnIdentifier();
			addExperimentBackgroundColumnIdentifier();
			addExperimentEnvelopeColumnIdentifier();
			addExperimentColumnIdentifier();
			separateExperiments();
		
			return true;
		}

	}

	return false;
		
}

void XPSExperimentImporter::setExperimentFileName()
{
	int index;
	string name = "";
	index = experimentFilePath.ReverseFind('\\');
	if (index != -1)
	{
		for(int i; i < experimentFilePath.GetLength(); i++)
		{
			string temporal = experimentFilePath[i];
			name += temporal;
		}
		experimentFileName = name;
	}
	else
		experimentFileName = "None";
}

string XPSExperimentImporter::getExperimentFileName()
{
	return experimentFileName;
}

void XPSExperimentImporter::fixColumns(Worksheet wks)
{	
    int numberOfColumns = wks.GetNumCols();
	for (int i = 0; i < numberOfColumns; i++)
	{
		Column col(wks, i);
		if (col.GetLongName() == "B.E.")
		{
			col.SetUnits("Binding energy, eV");
			col.SetType(3);
		}
		else if (col.GetLongName() == "")
		{
			wks.DeleteCol(i);
			i --;
			numberOfColumns -= 1;
		}
		else
			col.SetUnits("CPS");
	}

}


int XPSExperimentImporter::getNumberOfSamples()
{
	return numberOfSamples;
}

string XPSExperimentImporter::getNameBeforeChar(string nameOfColumn, char characterToSearch)
{
	string name = "";
	int index = nameOfColumn.ReverseFind(characterToSearch);
	for(int i = 0; i < index; i++)
	{
		string temporal = nameOfColumn[i];
		name += temporal;
	}
	return name;
}


void XPSExperimentImporter::setExperimentsNames()
{
	vector<string> addedNames();
	
	foreach(Column column in experimentWorksheet.Columns)
	{	
		if(columnHasExperimentName(column))
		{
			char experimentNameDelimiter = ':';
			string experimentName = getNameBeforeChar(column.GetLongName(), experimentNameDelimiter);
			
			if(columnNameIncludesCompound(experimentName))
				experimentName = getNameBeforeCompound(experimentName);
			
			if(experimentNameIsNotAlreadyAdded(addedNames, experimentName))
			{
				experimentsNames.Add(experimentName);
				addedNames.Add(experimentName);
			}
		}
		
	}
}

bool XPSExperimentImporter::experimentNameIsNotAlreadyAdded(vector<string> experimentsNames, string experimentName)
{
	for(int i = 0; i < experimentsNames.GetSize(); i++)
	   {
		   if(experimentName == experimentsNames[i])
		     {
				return false;  
			 }
        }
    
        return true;
}


bool XPSExperimentImporter::columnHasExperimentName(Column columnObject)
{
	if(     columnObject.GetLongName() != "B.E." 
		&&  columnObject.GetLongName() != "Background" 
		&&  columnObject.GetLongName() != "Envelope")
		return true;
	else
		return false;
}

vector<string> XPSExperimentImporter::getExperimentsNames()
{
	return experimentsNames;
}

int XPSExperimentImporter::getNumberOfExperiments()
{
	return experimentsNames.GetSize();
}

string XPSExperimentImporter::getNameBeforeCompound(string block)
{
	string name = "";
	int index = block.ReverseFind('_');

	for(int i = 0; i < index; i++)
	{
		string temporal = block[i];
		name += temporal;
	}	
	return name;
}

string XPSExperimentImporter::getCompound(string experimentName)
{
	string name = "";
	
	int startingIndex = experimentName.ReverseFind('_');	
	int finalIndex = experimentName.ReverseFind(':');
	
	if(startingIndex != -1 && finalIndex != -1)
	{
		startingIndex = startingIndex + 1;
	
		for(startingIndex; startingIndex < finalIndex; startingIndex++)
		 {
			string temp = experimentName[startingIndex];
			name += temp;
		 }	
	}
	
	return name;
}

void XPSExperimentImporter::separateExperiments()
{
	int amountOfExperiments = experimentsNames.GetSize();
	
	for(int index = 0; index < amountOfExperiments; index++)
	{
		WorksheetPage worksheetPage;
		worksheetPage.Create(experimentsNames[index], CREATE_HIDDEN);
		worksheetPage.SetName(experimentsNames[index]);
		worksheetPage.SetLongName(experimentsNames[index]);

		
		bindExperimentData(worksheetPage, experimentsNames[index]);
	}
}

void XPSExperimentImporter::bindExperimentData(WorksheetPage experimentWorksheetPage, string experimentName)
{
	string worksheetPageName = experimentWorksheetPage.GetName();
	
	Worksheet worksheet(worksheetPageName);
	
	int numberOfColumns = experimentWorksheet.GetNumCols();

	int layerIndex = 0;
	vector<string> addedLayers();
	
	
	foreach(Column column in experimentWorksheet.Columns)
	{
		string columnName = column.GetLongName();
	
		string longColumnName = columnName;
		
		string compound = getCompound(longColumnName);
	
		if(columnHasExperimentName(column))
			columnName = getNameBeforeChar(columnName, ':');
		
		if(columnNameIncludesCompound(columnName))
			columnName = getNameBeforeCompound(columnName);
		
		 if(column.GetComments() == experimentName)
		   {		
		         if(compound != "")
		         {
		         	if(experimentNameIsNotAlreadyAdded(addedLayers, compound))
		         	{
		         		if(longColumnName.Find(compound) > -1)
		         		{
		         			Layer layer = experimentWorksheetPage.Layers(layerIndex);
		         			
		         			if(layer == NULL)
								 layerIndex = experimentWorksheetPage.AddLayer();
		         			
							experimentWorksheetPage.Layers(layerIndex).SetName(compound);
							
							bindLayer(compound, experimentWorksheetPage.Layers(layerIndex), experimentName);
							
							addedLayers.Add(compound);
							
							layerIndex++;

		         		}
		         	}
		         }
		   }
	}

	
}

void XPSExperimentImporter::bindLayer(string layerName, Worksheet layerWorksheet, string experimentName)
{
	int layerWorksheetColumnIndex = 0;
	int numberOfColumns = experimentWorksheet.GetNumCols();
	
	for(int i = 0; i < numberOfColumns; i++)
	{
		Column column(experimentWorksheet, i);
		string columnName = column.GetLongName();
		
		if(columnHasExperimentName(column))
			columnName = getNameBeforeChar(columnName, ':');
		
		if(columnNameIncludesCompound(columnName))
			columnName = getNameBeforeCompound(columnName);
		
		if(column.GetComments() == experimentName)
		{
			if(columnName == "B.E.")
			{
				Column nextColumn(experimentWorksheet, i+1);
				string nextColumnName = nextColumn.GetLongName();
				
				if(nextColumnName.Find("_"+layerName) > - 1)
				{
					Column layerWorksheetColumn(layerWorksheet, layerWorksheetColumnIndex);
			     
					if(layerWorksheetColumn == NULL)
					  {
						layerWorksheetColumnIndex = layerWorksheet.AddCol();
						layerWorksheetColumn = layerWorksheet.Columns(layerWorksheetColumnIndex);
					  }
			     
						layerWorksheetColumn.SetType(OKDATAOBJ_DESIGNATION_X);
			     
						layerWorksheetColumn.SetLongName(column.GetLongName());
						layerWorksheetColumn.SetUnits(column.GetUnits());
			     
						Dataset layerWorkSheetColumnDataSource(column);
						Dataset layerWorkSheetColumnData(layerWorksheetColumn);
			     
						layerWorkSheetColumnData = layerWorkSheetColumnDataSource;
			    
			     
						layerWorksheetColumnIndex++;
				}
				
			}
			if(columnName == "Envelope")
			{
				Column previousColumn(experimentWorksheet, i-2);
				string previousColumnName = previousColumn.GetLongName();
				
				if(previousColumnName.Find("_"+layerName) > - 1)
				{
					Column layerWorksheetColumn(layerWorksheet, layerWorksheetColumnIndex);
			     
					if(layerWorksheetColumn == NULL)
					  {
						layerWorksheetColumnIndex = layerWorksheet.AddCol();
						layerWorksheetColumn = layerWorksheet.Columns(layerWorksheetColumnIndex);
					  }
			     
						layerWorksheetColumn.SetLongName(column.GetLongName());
						layerWorksheetColumn.SetUnits(column.GetUnits());
			     
						Dataset layerWorkSheetColumnDataSource(column);
						Dataset layerWorkSheetColumnData(layerWorksheetColumn);
			     
						layerWorkSheetColumnData = layerWorkSheetColumnDataSource;
			    
			     
						layerWorksheetColumnIndex++;
				}
				
			}
			if(columnName == "Background")
			{
				Column previousColumn(experimentWorksheet, i-1);
				string previousColumnName = previousColumn.GetLongName();
				
				if(previousColumnName.Find("_"+layerName) > - 1)
				{
					Column layerWorksheetColumn(layerWorksheet, layerWorksheetColumnIndex);
			     
					if(layerWorksheetColumn == NULL)
					  {
						layerWorksheetColumnIndex = layerWorksheet.AddCol();
						layerWorksheetColumn = layerWorksheet.Columns(layerWorksheetColumnIndex);
					  }
			     
						layerWorksheetColumn.SetLongName(column.GetLongName());
						layerWorksheetColumn.SetUnits(column.GetUnits());
			     
						Dataset layerWorkSheetColumnDataSource(column);
						Dataset layerWorkSheetColumnData(layerWorksheetColumn);
			     
						layerWorkSheetColumnData = layerWorkSheetColumnDataSource;
			    
			     
						layerWorksheetColumnIndex++;
				}
			}
			    
			if(column.GetLongName().Find("_"+layerName) > -1)
			{
				Column layerWorksheetColumn(layerWorksheet, layerWorksheetColumnIndex);
			     
			     if(layerWorksheetColumn == NULL)
			     {
			     	layerWorksheetColumnIndex = layerWorksheet.AddCol();
			     	layerWorksheetColumn = layerWorksheet.Columns(layerWorksheetColumnIndex);
			     }
			     
			     if(column.GetLongName() == "B.E.")
			     	layerWorksheetColumn.SetType(OKDATAOBJ_DESIGNATION_X);
			     
			     layerWorksheetColumn.SetLongName(column.GetLongName());
			     layerWorksheetColumn.SetUnits(column.GetUnits());
			     
			     Dataset layerWorkSheetColumnDataSource(column);
			     Dataset layerWorkSheetColumnData(layerWorksheetColumn);
			     
			     layerWorkSheetColumnData = layerWorkSheetColumnDataSource;
			    
			     
			     layerWorksheetColumnIndex++;
				
			}
		}
	}
}

void XPSExperimentImporter::addExperimentBEColumnIdentifier()
{
	int amountOfExperiments = experimentsNames.GetSize();
	int numberOfColumns = experimentWorksheet.GetNumCols();
	
	for(int index = 0; index < amountOfExperiments; index++)
	{
		for(int i = 0; i < numberOfColumns; i++)
		{
		   Column column(experimentWorksheet,i);
		   string columnName = column.GetLongName();
		   
		   if(columnName == "B.E.")
		   {
				Column nextColumn(experimentWorksheet, i+1);
				string nextColumnName = nextColumn.GetLongName();
				
			    if(columnHasExperimentName(nextColumn))
					nextColumnName = getNameBeforeChar(nextColumnName, ':');
		
				if(columnNameIncludesCompound(nextColumnName))
					nextColumnName = getNameBeforeCompound(nextColumnName);
				
				if(nextColumnName == experimentsNames[index] && column.GetComments()=="")
					column.SetComments(experimentsNames[index]);
				else if(nextColumnName != experimentsNames[index] && column.GetComments()=="")
				{
					column.SetComments(nextColumnName);	
				}
		   }		   
		}
	}
}

void XPSExperimentImporter::addExperimentBackgroundColumnIdentifier()
{
	int numberOfColumns = experimentWorksheet.GetNumCols();
	

		for(int i = 0; i < numberOfColumns; i++)
		{
		   Column column(experimentWorksheet,i);
		   string columnName = column.GetLongName();
		   
		   if(columnName == "Background")
		   {
				
			    Column previousColumn(experimentWorksheet, i-1);
			    string columnName = previousColumn.GetLongName();
					
			    if(columnHasExperimentName(previousColumn))
					columnName = getNameBeforeChar(columnName, ':');
		
				if(columnNameIncludesCompound(columnName))
					columnName = getNameBeforeCompound(columnName);
					
				if(column.GetComments()=="")
					column.SetComments(columnName);
		   }		   
		}
}

void XPSExperimentImporter::addExperimentEnvelopeColumnIdentifier()
{
		int numberOfColumns = experimentWorksheet.GetNumCols();
	

		for(int i = 0; i < numberOfColumns; i++)
		{
		   Column column(experimentWorksheet,i);
		   string columnName = column.GetLongName();
		   
		   if(columnName == "Envelope")
		   {
				
			    Column previousColumn(experimentWorksheet, i-1);
			    string previousColumnExperimentComment = previousColumn.GetComments();
					
				if(column.GetComments()=="")
					column.SetComments(previousColumnExperimentComment);
		   }		   
		}
	
}

void XPSExperimentImporter::addExperimentColumnIdentifier()
{
	int amountOfExperiments = experimentsNames.GetSize();
	int numberOfColumns = experimentWorksheet.GetNumCols();
	
	for(int index = 0; index < amountOfExperiments; index++)
	{
		for(int i = 0; i < numberOfColumns; i++)
		{
		   Column column(experimentWorksheet,i);
		   string columnName = column.GetLongName();
			
		   if(columnHasExperimentName(column))
			  columnName = getNameBeforeChar(columnName, ':');
		
		   if(columnNameIncludesCompound(columnName))
			  columnName = getNameBeforeCompound(columnName);
		   
		   if(columnName == "B.E." || columnName == "Envelope" || columnName == "Background" || columnName == experimentsNames[index])
		   {
		        if(column.GetComments()=="")
					column.SetComments(experimentsNames[index]);
		   }		   
		}
	}
}



bool XPSExperimentImporter::columnNameIncludesCompound(string columnName)
{
    bool included;
    
    if(columnName.Find("_") > -1)
    	return included = true;
    else
        return included = false;
}

