//
//  main.m
//  business-impact-map-assistant
//
//  Created by Tommy on 2014-02-24.
//  Copyright (c) 2014 Helt Enkelt AB. All rights reserved.
//

#import <Cocoa/Cocoa.h>

#import <AppleScriptObjC/AppleScriptObjC.h>

int main(int argc, const char * argv[])
{
    [[NSBundle mainBundle] loadAppleScriptObjectiveCScripts];
    return NSApplicationMain(argc, argv);
}
