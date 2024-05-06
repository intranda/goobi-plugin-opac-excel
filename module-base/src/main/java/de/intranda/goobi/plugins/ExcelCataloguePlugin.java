package de.intranda.goobi.plugins;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.reloading.FileChangedReloadingStrategy;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.goobi.production.enums.PluginType;
import org.goobi.production.plugin.interfaces.IOpacPlugin;

import de.intranda.goobi.plugins.util.GroupMappingObject;
import de.intranda.goobi.plugins.util.MetadataMappingObject;
import de.intranda.goobi.plugins.util.PersonMappingObject;
import de.sub.goobi.config.ConfigPlugins;
import de.unigoettingen.sub.search.opac.ConfigOpac;
import de.unigoettingen.sub.search.opac.ConfigOpacCatalogue;
import de.unigoettingen.sub.search.opac.ConfigOpacDoctype;
import lombok.Data;
import lombok.extern.log4j.Log4j;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.DocStructType;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.dl.MetadataGroup;
import ugh.dl.Person;
import ugh.dl.Prefs;
import ugh.exceptions.MetadataTypeNotAllowedException;
import ugh.exceptions.UGHException;
import ugh.fileformats.mets.MetsMods;

@PluginImplementation
@Log4j
@Data
public class ExcelCataloguePlugin implements IOpacPlugin {

    private PluginType type = PluginType.Opac;

    private String title = "intranda_opac_generic_excel";

    private String gattung = "Monograph";

    private int hitcount;

    private String atstsl = "";

    private ConfigOpacCatalogue coc;

    private int identifierColumnNumber;

    private String excelFileName;

    private List<MetadataMappingObject> metadataList = new ArrayList<>();
    private List<PersonMappingObject> personList = new ArrayList<>();
    private List<GroupMappingObject> groupList = new ArrayList<>();

    @SuppressWarnings("unchecked")
    public void loadConfig() {
        XMLConfiguration xmlConfig = ConfigPlugins.getPluginConfig(this);
        xmlConfig.setExpressionEngine(new XPathExpressionEngine());
        xmlConfig.setReloadingStrategy(new FileChangedReloadingStrategy());

        SubnodeConfiguration myconfig = null;
        try {
            myconfig = xmlConfig.configurationAt("//config");
            gattung = myconfig.getString("/publicationType", "Monograph");
            identifierColumnNumber = myconfig.getInt("identifierColumn", 0);
            excelFileName = myconfig.getString("filename");

            List<HierarchicalConfiguration> mml = xmlConfig.configurationsAt("//metadata");
            for (HierarchicalConfiguration md : mml) {
                metadataList.add(getMetadata(md));
            }

            List<HierarchicalConfiguration> pml = xmlConfig.configurationsAt("//person");
            for (HierarchicalConfiguration md : pml) {
                personList.add(getPersons(md));
            }

            List<HierarchicalConfiguration> gml = xmlConfig.configurationsAt("//group");
            for (HierarchicalConfiguration md : gml) {
                String rulesetName = md.getString("@ugh");
                GroupMappingObject grp = new GroupMappingObject();
                grp.setRulesetName(rulesetName);
                List<HierarchicalConfiguration> subList = md.configurationsAt("//person");
                for (HierarchicalConfiguration sub : subList) {
                    PersonMappingObject pmo = getPersons(sub);
                    grp.getPersonList().add(pmo);
                }

                subList = md.configurationsAt("//metadata");
                for (HierarchicalConfiguration sub : subList) {
                    MetadataMappingObject pmo = getMetadata(sub);
                    grp.getMetadataList().add(pmo);
                }

                groupList.add(grp);

            }

        } catch (IllegalArgumentException e) {
        }
    }

    @Override
    public Fileformat search(String inSuchfeld, String inSuchbegriff, ConfigOpacCatalogue coc, Prefs prefs) throws Exception {
        this.coc = coc;
        Fileformat fileformat = null;
        loadConfig();
        // read excel file
        InputStream file = null;
        try {
            file = new FileInputStream(excelFileName);

            BOMInputStream in = new BOMInputStream(file, false);

            Workbook wb = WorkbookFactory.create(in);

            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();
            Map<Integer, String> map = new HashMap<>();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                Cell identifierCell = row.getCell(identifierColumnNumber);
                if (identifierCell != null) {
                    String value = getCellValue(identifierCell);
                    // found correct row
                    if (StringUtils.isNotBlank(value) && value.equals(inSuchbegriff)) {
                        Iterator<Cell> cellIterator = row.cellIterator();
                        Integer i = 1;
                        while (cellIterator.hasNext()) {
                            Cell cell = cellIterator.next();
                            value = getCellValue(cell);
                            map.put(i, value);
                            i++;
                        }
                        break;
                    }
                }
            }
            if (map.isEmpty()) {
                // nothing found
                hitcount = 0;
                return null;
            }
            fileformat = generateMetadata(map, prefs);
        } catch ( IOException e) {
            log.error(e);

        } finally {
            if (file != null) {
                try {
                    file.close();
                } catch (IOException e) {
                    log.error(e);
                }
            }
        }

        return fileformat;
    }

    private Fileformat generateMetadata(Map<Integer, String> map, Prefs prefs) throws UGHException {
        Fileformat ff = new MetsMods(prefs);
        DigitalDocument digitalDocument = new DigitalDocument();
        ff.setDigitalDocument(digitalDocument);

        DocStructType logicalType = prefs.getDocStrctTypeByName(gattung);
        DocStruct logical = digitalDocument.createDocStruct(logicalType);
        digitalDocument.setLogicalDocStruct(logical);
        DocStructType physicalType = prefs.getDocStrctTypeByName("BoundBook");
        DocStruct physical = digitalDocument.createDocStruct(physicalType);
        digitalDocument.setPhysicalDocStruct(physical);
        Metadata imagePath = new Metadata(prefs.getMetadataTypeByName("pathimagefiles"));
        imagePath.setValue("./images/");
        physical.addMetadata(imagePath);
        for (MetadataMappingObject mmo : metadataList) {
            String value = map.get(mmo.getExcelColumn());
            String identifier = null;
            if (mmo.getIdentifierColumn() != null) {
                identifier = map.get(mmo.getIdentifierColumn());
            }
            if (StringUtils.isNotBlank(mmo.getRulesetName())) {
                try {
                    Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                    md.setValue(value);
                    if (identifier != null) {
                        // uri?
                        md.setAuthorityValue(identifier);
                    }
                    logical.addMetadata(md);
                } catch (MetadataTypeNotAllowedException e) {
                    log.info(e);
                    // Metadata is not known or not allowed
                }
            }
        }

        for (PersonMappingObject mmo : personList) {
            String firstname = map.get(mmo.getFirstnameColumn());
            String lastname = map.get(mmo.getLastnameColumn());
            String identifier = null;
            if (mmo.getIdentifierColumn() != null) {
                identifier = map.get(mmo.getIdentifierColumn());
            }
            if (StringUtils.isNotBlank(mmo.getRulesetName())) {
                try {
                    Person p = new Person(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                    p.setFirstname(firstname);
                    p.setLastname(lastname);

                    if (identifier != null) {
                        // TODO uri?
                        p.setAuthorityValue(identifier);
                    }
                    logical.addPerson(p);
                } catch (MetadataTypeNotAllowedException e) {
                    log.info(e);
                    // Metadata is not known or not allowed
                }
            }
        }

        for (GroupMappingObject gmo : groupList) {
            try {
                MetadataGroup group = new MetadataGroup(prefs.getMetadataGroupTypeByName(gmo.getRulesetName()));
                for (MetadataMappingObject mmo : gmo.getMetadataList()) {
                    String value = map.get(mmo.getExcelColumn());
                    Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                    md.setValue(value);
                    if (mmo.getIdentifierColumn() != null) {
                        md.setAuthorityValue(map.get(mmo.getIdentifierColumn()));
                    }
                    group.addMetadata(md);
                }
                for (PersonMappingObject pmo : gmo.getPersonList()) {
                    Person p = new Person(prefs.getMetadataTypeByName(pmo.getRulesetName()));
                    p.setFirstname(map.get(pmo.getFirstnameColumn()));
                    p.setLastname(map.get(pmo.getLastnameColumn()));

                    if (pmo.getIdentifierColumn() != null) {
                        // TODO uri?
                        p.setAuthorityValue(map.get(pmo.getIdentifierColumn()));
                    }
                    group.addMetadata(p);
                }
                logical.addMetadataGroup(group);

            } catch (MetadataTypeNotAllowedException e) {
                log.info(e);
                // Metadata is not known or not allowed
            }
        }
        return ff;
    }

    private String getCellValue(Cell cell) {
        String value = null;
        switch (cell.getCellType()) {
            case BOOLEAN:
                value = cell.getBooleanCellValue() ? "true" : "false";
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case NUMERIC:
                value = String.valueOf((int) cell.getNumericCellValue());
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            default:
                // none, error, blank
                value = "";
                break;
        }
        return value;
    }

    @Override
    public String createAtstsl(String value, String value2) {
        // TODO Auto-generated method stub
        return "";
    }

    @Override
    public ConfigOpacDoctype getOpacDocType() {
        ConfigOpac co = ConfigOpac.getInstance();
        ConfigOpacDoctype cod = co.getDoctypeByMapping(this.gattung, this.coc.getTitle());
        if (cod == null) {
            cod = co.getAllDoctypes().get(0);
            this.gattung = cod.getMappings().get(0);

        }
        return cod;
    }

    private MetadataMappingObject getMetadata(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        String propertyName = md.getString("@name");
        Integer columnNumber = md.getInteger("@column", null);
        Integer identifierColumn = md.getInteger("@identifier", null);
        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setExcelColumn(columnNumber);
        mmo.setIdentifierColumn(identifierColumn);
        mmo.setPropertyName(propertyName);
        mmo.setRulesetName(rulesetName);
        return mmo;
    }

    private PersonMappingObject getPersons(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        Integer firstname = md.getInteger("firstname", null);
        Integer lastname = md.getInteger("lastname", null);
        Integer identifier = md.getInteger("identifier", null);
        PersonMappingObject pmo = new PersonMappingObject();
        pmo.setFirstnameColumn(firstname);
        pmo.setLastnameColumn(lastname);
        pmo.setIdentifierColumn(identifier);
        pmo.setRulesetName(rulesetName);
        return pmo;

    }
}
