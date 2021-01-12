import java.util.ArrayList;
import java.util.Collections;
import java.util.Scanner;

class Pet {
  public static void main(String args[]){
    Scanner reader=new Scanner(System.in);
    ArrayList<Integer> arr=new ArrayList<Integer>();
    ArrayList<Integer> lst=new ArrayList<Integer>();
    for(int i=0;i<10;i++){
      int s=reader.nextInt();
      arr.add(s);
    }
    int count=0;
    int loc=9;
    for(int i=0;i<10;i++){
      for(int j=loc;j>=1;j--){
        count++;
        if(arr.get(j)!=arr.get(j-1)){
          loc=j-1;
          break;
        }

      }
      lst.add(count);
      count=0;
    }
    if(arr.get(0)==arr.get(1)){
      lst.set(arr.size()-1,lst.get(lst.size()-1)+1);
    }
    else{
      lst.add(1);
    }
    int max=Collections.max(lst);
    System.out.println(arr.size());
    System.out.println(lst.size());
    System.out.println(max);
    System.out.println(arr);
    System.out.println(lst);

  }


}