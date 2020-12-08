pipeline {
  agent any
  stages {
    stage('Call AD COM server') {
      steps {
        sh '''#!/bin/bash

wget -q -O API_AddLib.py https://github.com/laomudashabi/ECS-Verification/blob/main/API_AddLib.py 
/usr/bin/python API_AddLib.py $"C:\\\\Users\\\\a269028\\\\Desktop\\\\TestScope.xml" "C:\\\\Program Files (x86)\\\\CVDE-Interface\\\\Template_TestExecutionProject_AD5-6.zip" $"C:\\\\Users\\\\a269028\\\\Desktop\\\\"'''
      }
    }

  }
}