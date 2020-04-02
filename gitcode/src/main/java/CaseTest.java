import java.util.Scanner;

public class CaseTest {
    public static void main(String[] args){
        Scanner sc = new Scanner(System.in);
        int num = sc.nextInt();
        int sum = sc.nextInt();
        int[] value = new int[num];
        for(int i = 0; i < num; i++){
            value[i] = sc.nextInt();
        }
        int number = maxNumber(value,sum);
        System.out.println(number);
    }

    private static int maxNumber(int[] value, int sum) {
        int SumValue = 0,number;
        for(int i : value){
            SumValue += i;
        }
        if(SumValue < sum){
            number = (sum/SumValue)*value.length + mNumber(value,sum % SumValue);
        }else if(SumValue > sum){
            number = mNumber(value,sum);
        }else{
            number = value.length;
        }
        return number;
    }



    private static int mNumber(int[] value, int sum) {
        for(int i = 0 ;i < value.length ;i++){
            int number = 0;
            for(int j = 0 ;j < value.length ;j++){
                  if(sum < value[j]){
                      continue;
                  }else if(sum > value[j]){
                      sum -= value[j];
                      number++;
                  }else {
                      number++;
                      break;
                  }
            }
        }

        return 1;
    }

}
