pipeline {
  agent any
  stages {
    stage('Call AD COM server') {
      steps {
        sh '''#!/bin/bash

wget -q -O API_AddLib.py https://github.com/laomudashabi/ECS-Verification/blob/main/API_AddLib.py 
/usr/bin/python API_AddLib.py ${ARG1} ${ARG2} ${ARG3}'''
      }
    }

  }
}