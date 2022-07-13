export class MynewcustomlibrarydemoLibrary {
  public name(): string {
    return 'MynewcustomlibrarydemoLibrary';
  }

  public getCurrentTime(): string {
    
    let currentDate: Date;
    let str: string;

    currentDate = new Date();

    str = "<br>Todays Date is: " + currentDate.toDateString();
    str += "<br>Current Time is: " + currentDate.toTimeString();

    return(str);
  }
}
