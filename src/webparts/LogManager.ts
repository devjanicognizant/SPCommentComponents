import pnp  from 'sp-pnp-js';
import {
  Logger,
  ConsoleListener,
  LogLevel
} from "sp-pnp-js";

// Subscribe a console listener to enable logging mechanism
Logger.subscribe(new ConsoleListener());

// Set the active log level as information
Logger.activeLogLevel = LogLevel.Error;

// Helper class to manage exception, log etc
export default class LogManager
{
    // Log exception along with component name and method name from where the exception is generated
    public static logException(exception:any, msg:string, componentName:string, methodName: string){
        Logger.write("Component Name: " + componentName
                    +". Method Name: " + methodName
                    +". Message:" + msg
                    +". Error Message:" + exception.message, LogLevel.Error);
        //Logger.write(exception, LogLevel.Error);
    }

    // Log any mesage
    public static logMessage(msg:string){
        Logger.write(msg, LogLevel.Error);
    }

    // Convert the object to json and log the same
    public static logObject(obj:any){
        Logger.writeJSON(JSON.stringify(obj), LogLevel.Error);
    }
}