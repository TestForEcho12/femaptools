import sys
import pyfemap
import pythoncom
import numpy as np
import pandas as pd

class FemapTools:
    
    def __init__(self):
        try: 
            existObj = pythoncom.connect(pyfemap.model.CLSID)
            self.app = pyfemap.model(existObj)
        except:
            sys.exit('Femap is not open')   
            
    def get_app(self):
        return self.app
    
    def message(self, color, message):
        '''
        0 = Normal
        1 = Highlight
        2 = Warning
        3 = Error
        '''
        self.app.feAppMessage(color, message)
    
    def _create_set(self, entity_id, title, entity_list=None, select_all=False):
        entitySet = self.app.feSet
        if entity_list:
            _ = entitySet.AddArray(len(entity_list), entity_list)
        elif select_all:
            _ = entitySet.AddAll(entity_id)
        else:
            _ = entitySet.Select(entity_id, True, title)
        return entitySet
            
    def create_set_of_nodes(self, node_list=None, select_all=False):
        return self._create_set(7, 'Select Nodes', entity_list=node_list, select_all=select_all)
    
    def create_set_of_elements(self, element_list=None, select_all=False):
        return self._create_set(8, 'Select Elements', entity_list=element_list, select_all=select_all)
    
    def create_set_of_materials(self, material_list=None, select_all=False):
        return self._create_set(10, 'Select Materials', entity_list=material_list, select_all=select_all)
    
    def create_set_of_properties(self, property_list=None, select_all=False):
        return self._create_set(11, 'Select Properties', entity_list=property_list, select_all=select_all)
    
    def create_set_of_outputs(self, select_all=False):
        outputSet = self.app.feSet
        if select_all:
            _ = outputSet.AddAll(28)
        else:
            _ = outputSet.SelectMultiIDV2(28, 1, 'Select Output Sets')
        return outputSet
    
    def get_list_from_Femap_set(self, feSet):
        [_, _, entityList] = feSet.GetArray()
        return entityList
    
    def _get_results(self, outputSet, vectors, entitySet, entityTypeID, **kwargs):
        '''
        Parameters
        ----------
        outputSet : Femap Set
            Includes the output set ids.
        vectors : List
            Contains the output vector ids.
        entitySet : Femap Set
            Includes the entity ids.
        entityTypeID : Integer
            Femap API Entity Type (Reference Section 3.3.6 of Femap API 
            Documentation).
        **transform : String, optional
                - 'Nodal': transform nodal results to the node's output CSys

        Returns
        -------
        pandas dataframe
            Dataframe containing the results. Columns are:
                - 'id': Entity ID
                - 'set': output set
                - [vector id]: columns are the output vector ids
        '''
        transform = kwargs.get('transform')
        results = []
        output = self.app.feResults
        [_, _, outputIDs] = outputSet.GetArray()
        entitySetID = entitySet.ID
        [_, nEntities, entityIDs] = entitySet.GetArray()
        nVectors = len(vectors)
        columnIndex = [i for i in range(nVectors)]
        
        for set_id in outputIDs:
            _ = output.clear()
            _ = output.DataNeeded(entityTypeID, entitySetID)
            
            for vect in vectors:
                [_, _, _] = output.AddColumnV2(set_id, vect, False)
            if transform == 'nodal':
                _ = output.SetNodalTransform(2, 0)
            _ = output.Populate()   
            [_, dVals, _] = output.GetRowsAndColumnsByID(entitySetID, nVectors, columnIndex)
            dVals = np.reshape(dVals, [nEntities, nVectors])
            df = pd.concat([pd.DataFrame({'id': entityIDs, 'set': set_id}), 
                    pd.DataFrame(dVals, columns=vectors)], 
                    axis=1)     
            results.append(df)
        return pd.concat(results, ignore_index=True)
    
    def get_element_results(self, outputSet, elementSet, vectors):
        '''
        outputSet = Femap set including output set ids
        elementSet = Femap set including element ids
        vectors = List of output vector ids
        '''
        return self._get_results(outputSet, vectors, elementSet, 8)
            
    def get_node_results(self, outputSet, nodeSet, vectors, transform=False):
        '''
        outputSet = Femap set including output set ids
        nodeSet = Femap set including node ids
        vectors = List of output vector ids
        transform = Boolian to transform results to output CSys
        '''
        tsfm = 'nodal' if transform else ''
        return self._get_results(outputSet, vectors, nodeSet, 7, transform=tsfm)
    
    def get_dict_of_properties_from_element_set(self, elementSet):
        element = self.app.feElem
        [_, _, entID, propID, _, _, _, _, _, _, _, _, _, _, _, _, _] = element.GetAllArray(elementSet.ID)
        propIDs = dict((e, p) for e, p in zip(entID, propID))
        return propIDs
    
    def get_dict_of_output_titles_from_output_set(self, outputSet):
        [_, _, outputIDs] = outputSet.GetArray()
        output = self.app.feResults
        titles = {}
        for set_id in outputIDs:
            [_, title] = output.SetTitle(set_id)
            titles[set_id] = title
        return titles
            
    def get_dict_of_frequencies_from_output_set(self, outputSet):
        [_, _, outputIDs] = outputSet.GetArray()
        output = self.app.feResults
        frequencies = {}
        for setID in outputIDs:
            [_, _, _, dSetValue] = output.SetInfo(setID)
            frequencies[setID] = dSetValue
        return frequencies
    
    def delete_output(self):
        self.app.feDeleteAll(False, False, True, True)
    
    def import_output(self, output_path):
        self.app.feFileReadNastranResults(0, output_path)
