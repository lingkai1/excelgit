class Student(object):
    __slots__=('name','age')
    
class Student(object):
    def get_score(self):
        return self._score
    score = 1000
    def set_score(self,value):
        if not isinstance(value,int):
            raise ValueError('score must be a interge')
        if value < 0 or value > 100:
            raise ValueError('score must between 0~ 100')
        self._score = value
    def __str__(self):
        return "student score (name: %s)" %self.score
    __repr__=__str__

        

