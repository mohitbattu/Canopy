CODE_CHANGES = getGitChanges()
def gv
pipeline{
  agent any
  parameters{
  choice(name: 'VERSION',choices: ['1.1.0','1.2.0','1.3.0'],description: '')
  booleanParam(name: 'executeTests',defaultValue: true,description:'')
  }
  environment{
  NEW_VERSION='1.3.0'
  SERVER_CREDENTIALS=credentials('server-credentials')
  }
  tools{
  gradle 'Gradle-6.7'
  }
  stages{
    stage("init") {
            steps {
                script {
                   gv = load "script.groovy" 
                }
            }
        }

    stage('build'){
      steps{
        when{
          expression{
          BRANCH_NAME == 'dev' && CODE_CHANGES == true
          }
        }
        gv.buildApp()
        echo "building the application..${NEW_VERSION}"
        nodejs('Node-10.7'){
          sh 'yarn install'
          
        }
        echo "Application built"
      }
    }
    stage('test'){
      steps{
        when{
          expression{
          BRANCH_NAME == 'dev' || env.BRANCH_NAME == 'master'
          }
        }
        echo "testing the application.."
        echo "Tested the Application"
       
          sh './gradlew -v'
       
      }
    }
    stage('deploy'){
      steps{
        echo "deploying the application.."
        echo "Deployed the Application"
        withCredentials([
        usernamePassword(credentials: 'server-credentials',usernameVariable:USER,passwordVariable: PWD)
        ]){
          sh "some script ${USER} ${PWD}" 
        }
      }
    }
  }
  post{
    always{
    //
    }
    success{
    //
    }
  }
}
