/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package formulacoding;
import java.lang.Math;
import java.io.IOException;
import javax.swing.JFrame;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowListener;
import java.awt.event.*;
import java.util.Date;
import java.util.Vector;
import javax.swing.JOptionPane;
import proj1.*;
import format.*;


/**
 *
 * @author Nisu
 */
/*class winclose extends WindowAdapter
{
    public void windowClosing(WindowEvent we)
    {
        System.exit(0);
    }
}*/
public class Formulacoding {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
     //   JFrame f=new JFrame("Input Page");
        int i=0,j=0;
        Proj1 p=new Proj1();
        Vector input=new Vector();
        Vector result=new Vector();
        Proj1.main(input);
        //winclose wc=new winclose();
  //  f.setSize(800, 600);
   // f.setVisible(true);
        String [] operator=new String[10];
        double siteID [];
        siteID=new double[10];
        String [] name=new String[10];
        double dt_of_cmsning [];
        dt_of_cmsning=new double[10];
        String [] adrss=new String[10];
        String [] lat_lon=new String[10];
        String [] rtt_gbt=new String[10];
        double [] bldng_ht=new double[10];
        double [] ant_ht=new double[10];
        String [] sys_type=new String[10];
        double [] bcf=new double[10];
        double [] car=new double[10];
        String [] mk_model=new String[10];
        double [] tot_tilt=new double[10];
        double [] ant_gain=new double[10];
        double [] ver_bw=new double[10];
        double [] sla=new double[10];
        double [] tx_power=new double[10];
        double [] com_loss=new double[10];
        double [] cab_len=new double[10];
        double [] unit_loss=new double[10];
        double [] EIRP=new double[10];
        double [] DTX_fac=new double[10];
        double [] ATPC_fac =new double[10];
        double [] EIRP_TCH=new double[10];
        double [] EIRP_tot=new double[10];
        
        int size=input.size();
        while(j<size)
        {
            operator[i]=((String)input.elementAt(j++));
            siteID[i]=((Double)input.elementAt(j++)).doubleValue();
            name[i]=((String)input.elementAt(j++));
            dt_of_cmsning[i]=((Double)input.elementAt(j++)).doubleValue();
            adrss[i]=((String)input.elementAt(j++));
            lat_lon[i]=((String)input.elementAt(j++));
            rtt_gbt[i]=((String)input.elementAt(j++));
            bldng_ht[i]=((Double)input.elementAt(j++)).doubleValue();
            ant_ht[i]=((Double)input.elementAt(j++)).doubleValue();
            sys_type[i]=((String)input.elementAt(j++));
            bcf[i]=((Double)input.elementAt(j++)).doubleValue();
            car[i]=((Double)input.elementAt(j++)).doubleValue();
            mk_model[i]=((String)input.elementAt(j++));
            tot_tilt[i]=((Double)input.elementAt(j++)).doubleValue();
            ant_gain[i]=((Double)input.elementAt(j++)).doubleValue();
            ver_bw[i]=((Double)input.elementAt(j++)).doubleValue();
            sla[i]=((Double)input.elementAt(j++)).doubleValue();
            tx_power[i]=((Double)input.elementAt(j++)).doubleValue();
            com_loss[i]=((Double)input.elementAt(j++)).doubleValue();
            cab_len[i]=((Double)input.elementAt(j++)).doubleValue();
            unit_loss[i]=((Double)input.elementAt(j++)).doubleValue();
             
            if("GSM".equals(sys_type[i]))
            {
                DTX_fac[i]=0.9;
                ATPC_fac[i]=0.9;
            }
            else if("CDMA".equals(sys_type[i]))
            {
                DTX_fac[i]=1;
                ATPC_fac[i]=1;
            }
            
            EIRP[i]=tx_power[i]-com_loss[i]-(cab_len[i]*(unit_loss[i])/100)+ant_gain[i];
          EIRP[i]=Math.pow(10,EIRP[i]/10)/1000;
    
            EIRP_tot[i]=EIRP[i]+(EIRP[i]*DTX_fac[i]*DTX_fac[i]*(car[i]-1));
        
            EIRP_TCH[i]=(EIRP[i]*ATPC_fac[i]*DTX_fac[i])*(2);
            
            result.addElement(new Double(EIRP[i]));
            result.addElement(new Double(DTX_fac[i]));
            result.addElement(new Double(ATPC_fac[i]));
            result.addElement(new Double(EIRP_TCH[i]));
            result.addElement(new Double(EIRP_tot[i]));
        }
        //JOptionPane.showMessageDialog(null, "System Type::"+sys_type);
        
        
    
      //  JOptionPane.showMessageDialog(null,"Operator::"+operator+"site ID::"+siteID+"Name"+name+"date"+dt_of_cmsning+"address"+adrss+"Lat_lon"+lat_lon+"rtt/gbt::"+rtt_gbt+""+"Value of bcf::"+bcf+"\ncar::"+car+"\ntot_tilt::"+tot_tilt+"\nant_gain::"+ant_gain+"\nver_bw::"+ver_bw+"\nsla::"+sla+"\ntx_power::"+tx_power+"\ncom loss::"+com_loss+"\ncab_len"+cab_len+"\nunit loss::"+unit_loss);
 
     //   JOptionPane.showMessageDialog(null,"DTX FACTOR::"+DTX_fac+"\nATPC_fac::"+ATPC_fac);
        
        
       // JOptionPane.showMessageDialog(null,"Value of EIRP::"+EIRP+"watts"+"\nValue of EIRP Total::"+EIRP_tot+"watts"+"EIRP_TCH::"+EIRP_TCH+"watts");
        
        
        format.main(input, result);
    }
}
