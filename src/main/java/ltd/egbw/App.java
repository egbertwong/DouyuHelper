package ltd.egbw;

import HTTPservice.getInfo;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        String roomInfo;
        getInfo getinfo = new getInfo();
        try {
            roomInfo = getinfo.run("1659672");
            System.out.println(roomInfo);
        } catch (Exception ignore){
            // nothing to do
        }
    }
}
