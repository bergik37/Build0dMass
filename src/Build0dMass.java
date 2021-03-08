import nxopen.*;
import nxopen.cae.ElementCreateBuilder;
import nxopen.cae.FEModel;
import nxopen.cae.FemPart;
import nxopen.cae.PropertyTable;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Iterator;


public class Build0dMass {

    public static void main(String[] args) throws Exception {

        int a=0;

        ArrayList <Double>arrX=new ArrayList();
        ArrayList <Double>arrY=new ArrayList();
        ArrayList <Double>arrZ=new ArrayList();
        ArrayList <Double>mass=new ArrayList();

        FileInputStream fileInputStream = new FileInputStream("D:\\координаты и массы.xls");
        Workbook wb = new HSSFWorkbook(fileInputStream);
        Sheet sheet = wb.getSheetAt(0);
        //DataFormatter formatter = new DataFormatter();
        Iterator<Row> rows = sheet.rowIterator();
        HSSFRow row;

        while (rows.hasNext()) {
            row = (HSSFRow) rows.next();
            arrX.add(row.getCell(0).getNumericCellValue());
            arrY.add(row.getCell(1).getNumericCellValue());
            arrZ.add(row.getCell(2).getNumericCellValue());
            mass.add(row.getCell(3).getNumericCellValue());
        }
        Session theSession = null;
        UFSession theUFSession = null;
        try {
            theSession = (Session) SessionFactory.get("Session");
            theUFSession = (UFSession) SessionFactory.get("UFSession");
            theUFSession.ui().openListingWindow();
            theUFSession.ui().writeListingWindow("Session and UFSession are created");
            /* Fill out the data structure */
            FemPart workFemPart = ((nxopen.cae.FemPart)theSession.parts().baseWork());
            ArrayList<Point3d> coordinates = new ArrayList<>();
            ElementCreateBuilder elementCreateBuilder1;
            FEModel fEModel1 = ((nxopen.cae.FEModel)workFemPart.findObject("FEModel"));
            NXObject nxObject1;
            nxopen.Point point1,point2;
            boolean added1;
            PropertyTable propertyTable1;
            Unit unit1 = ((nxopen.Unit)workFemPart.unitCollection().findObject("Kilogram"));
            for (int i = 0; i < arrX.size(); i++) {
                String telo=String.format("0d_manual_mesh(%d)",i+1);
                String number=String.format("ENTITY 2 %d 1",i+1);
                Point3d coordinates1 = new nxopen.Point3d();
                coordinates1.x = (arrX.get(i));
                coordinates1.y = (arrY.get(i));
                coordinates1.z = (arrZ.get(i));
                elementCreateBuilder1 = fEModel1.nodeElementMgr().createElementCreateBuilder();
                point1 = workFemPart.points().createPoint(coordinates1);
                point2 = workFemPart.points().createPoint(point1, null, nxopen.SmartObject.UpdateOption.AFTER_MODELING);
                added1 = elementCreateBuilder1.point().add(point2);
                elementCreateBuilder1.setElementDimensionOption(ElementCreateBuilder.ElemDimType.POINT);
                elementCreateBuilder1.elementType().setElementTypeName("CONM2");
                elementCreateBuilder1.setMeshName(telo);
                propertyTable1 = elementCreateBuilder1.elementType().propertyTable();
                propertyTable1.setBaseScalarWithDataPropertyValue("mass", mass.get(i), unit1);
                nxObject1 = elementCreateBuilder1.commit();

            }
            theUFSession.ui().writeListingWindow("\nCreated successfully");

        } catch (Exception ex) {
            if (theUFSession != null) {
                StringWriter s = new StringWriter();
                PrintWriter p = new PrintWriter(s);
                p.println("Caught exception " + ex);
                ex.printStackTrace(p);
                theUFSession.ui().writeListingWindow("\nFailed");
                //theUFSession.ui().writeListingWindow("\n"+ex.getMessage());
                theUFSession.ui().writeListingWindow("\n" + s.getBuffer().toString());
            }
        }
    }
}

